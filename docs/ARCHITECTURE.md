# Архитектура vba_engine

Цель: встраиваемая Rust-библиотека, которая исполняет код, совместимый с VBA
(Visual Basic for Applications), для автоматизации офисных приложений, и даёт
отладчик уровня VBE (шаги, брейкпоинты, watch, Immediate window).

## Слои

```
            ┌──────────────┐     ┌─────────────────────────────┐
 .bas/.cls  │  vba-syntax  │ AST │         vba-runtime         │
 ─────────► │ lexer→parser ├────►│ Interpreter ─ Variant/Value │
            └──────────────┘     │     │  │        stdlib      │
                                 │     │  └── trait Host ◄─────┼── офисное приложение
                                 │     │      trait Object ◄───┼── (Excel/Word-подобная
                                 │     └── trait DebugHook     │    объектная модель)
                                 └───────────┬─────────────────┘
                                             │
                              ┌──────────────▼───────────────┐
                              │          vba-debug           │
                              │ Debugger: режимы, breakpoints│
                              │ watches, Frontend / Channel  │
                              └──────────────┬───────────────┘
                                             │
                                ┌────────────▼────────────┐
                                │ vba-cli  (vba run/debug)│  будущее: DAP-сервер, IDE
                                └─────────────────────────┘
```

| Крейт | Ответственность |
|---|---|
| `vba-syntax` | Токенизатор (регистронезависимый, `_`-переносы, `'`/`Rem`, `""`, `&H`, `#дата#`, суффиксы типов), рекурсивный парсер, AST со `Span` на каждом операторе. `parse_expression` для watch-выражений. |
| `vba-runtime` | `Value` (Variant), правила приведения и операторов, `Object`/`Host` — интерфейс объектной модели, стандартная библиотека VBA, интерпретатор, хуки отладчика. |
| `vba-debug` | `Debugger` поверх `DebugHook`: режимы исполнения, брейкпоинты, watch, перехват ошибок, `ChannelFrontend` для управления из другого потока. |
| `vba-cli` | `vba run Main.bas Class1.cls ...`, `vba debug ...` — консольный отладчик с командами в духе VBE, брейкпоинты вида `Module:line`. |

## Система типов (`value.rs`)

`Value` = VBA Variant: `Empty, Null, Boolean, Byte, Integer(i16), Long(i32),
LongLong(i64), Single, Double, Currency(i64 ×10⁴), Date(f64 OLE),
String, Object(Option<ObjectRef>) (None = Nothing), Array, Error, Missing`.

- Объявленный тип (`Dim x As Long`) хранится в слоте переменной; присваивание
  приводит значение через `VarType::coerce` (банковское округление, Overflow = ошибка 6).
- Арифметика Variant расширяет тип при переполнении (Integer → Long → LongLong → Double),
  `+` на двух строках — конкатенация, `/` всегда Double, `\` и `Mod` — целочисленные.
- `Null` распространяется через операторы, в `If` считается False.
- Массивы — значения (копируются при присваивании), многомерные, с произвольными
  нижними границами, `ReDim Preserve`.
- `CStr(Double)` — 15 значащих цифр как в VBA (`0.1 + 0.2` → `0.3`).

## Объектная модель хоста (`object.rs`, `host.rs`)

Ключевая точка расширения — аналог `IDispatch`:

```rust
pub trait Object {
    fn type_name(&self) -> &str;
    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value>;
    fn default_value(&self) -> VbaResult<Value>;     // член по умолчанию ("")
    fn enumerate(&self) -> VbaResult<Vec<Value>>;     // For Each (_NewEnum)
    fn inspect(&self) -> Vec<(String, Value)>;        // окно Locals/Watch
}
pub enum Invoke { Get, Method, Let(Value), Set(Value) }
```

`Host` даёт глобальное пространство приложения: `global("ActiveSheet")`,
`global_object()` (неквалифицированные `Range("A1")` → `Application.Range`),
`debug_print`, `msgbox`, `inputbox`, `create_object`, `new_object`, `now`,
`poll` (DoEvents / отмена из UI). Встроены `Collection`, `Err`, `DynObject`
(объект-«мешок свойств» для быстрых моков и тестов).

Позднее связывание как в VBA: `Foo(1)` во время исполнения разрешается в
индекс массива, вызов процедуры, встроенную функцию, член по умолчанию объекта
или глобальный член хоста.

## Интерпретатор (`interp.rs`)

Tree-walking по AST. Фрейм = процедура + локальные слоты + текущая строка +
стек `With` + состояние `On Error`.

- Слоты `Rc<RefCell<Variable>>` — так реализован `ByRef` по умолчанию (аргумент
  делит слот с переменной вызывающего), `ByVal` копирует.
- Возвращаемое значение функции — локальный слот с её именем; `Fn(x)` внутри
  `Fn` с аргументами — рекурсивный вызов.
- Управление потоком: `Flow::{Exit, GoTo, Resume}` поднимается по вложенным
  блокам до блока, где есть метка — так работают `GoTo` и `On Error GoTo` из
  вложенных циклов.
- `On Error Resume Next` — пропуск оператора в самом внутреннем блоке;
  `On Error GoTo label` + `Resume` / `Resume Next` / `Resume label`; `Err.Raise`,
  `Err.Clear`; ошибки несут номер VBA и строку.
- `Option Explicit`, `Option Compare Text`, неявные переменные без Explicit.
- Лимит глубины рекурсии (ошибка 28 «Out of stack space»).

## Отладка

### Хук в рантайме (`debug.rs`)

```rust
pub trait DebugHook {
    fn on_statement(&mut self, ctx: &mut DebugContext) -> DebugAction; // перед каждым оператором и на End Sub
    fn on_enter / on_exit(...);                                         // вход/выход из процедуры
    fn on_stop(&mut self, ctx, StopKind) -> DebugAction;                // Stop, Debug.Assert
    fn on_error(&mut self, ctx, &VbaError, handled: bool) -> DebugAction;
}
```

`DebugContext` — живой доступ к остановленной программе: `location`,
`call_stack`, `locals(frame)`, `module_variables`, `eval_in_frame(frame, expr)`
(без побочных эффектов: неизвестные имена — ошибка, а не неявная переменная),
`execute(stmt)` (Immediate window: `x = 10`, `Debug.Print ...`).
Без установленного хука накладные расходы — одна проверка `Option` на оператор.

### Режимы `vba-debug`

| Режим | VBE | API |
|---|---|---|
| Обычный запуск | F5 без брейкпоинтов | без хука |
| Запуск под отладчиком | F5 | `Command::Continue` |
| Шаг с заходом | F8 | `Command::StepInto` |
| Шаг с обходом | Shift+F8 | `Command::StepOver` |
| Шаг с выходом | Ctrl+Shift+F8 | `Command::StepOut` |
| До курсора | Ctrl+F8 | `Command::RunToLine` |
| Остановка | Ctrl+Break | `Debugger::pause_handle()` (`Arc<AtomicBool>`) |
| Сброс | Reset | `Command::Abort` |
| Остановка на старте | F8 на макросе | `Debugger::stop_on_entry()` |

Брейкпоинты: по строке (и модулю), условные (`i = 2`), со счётчиком попаданий
(`Equals/AtLeast/Multiple`), logpoint-ы (`"i={i}"` печатается без остановки).

Watch: `Expression`, `BreakWhenTrue`, `BreakWhenChanged`, с ограничением контекста процедурой.

Ошибки: `ErrorBreak::{Never, Unhandled (по умолчанию, как в VBE), All}`.

Фронтенды: `Frontend` (в том же потоке, напр. консоль), `FnFrontend` (замыкание,
для тестов), `ChannelFrontend` — события `PauseEvent` (только `Send`-данные) и
запросы `Request` через `mpsc`, чтобы IDE/DAP-сервер управлял отладкой из
другого потока.

## Что уже работает

`Sub/Function` с `ByVal/ByRef/Optional/default`, именованные аргументы `x:=`,
`Dim/Const/ReDim [Preserve]/Static`, `If` (блочный и однострочный), `For/Next`,
`For Each`, `Do/Loop While/Until`, `While/Wend`, `Select Case` (`Is`, `To`, списки),
`With`, `Exit`, `GoTo`, `On Error`, `Resume`, `End`, `Stop`, `Debug.Print/Assert`,
`Set`, `New`, `Is`, `Like`, все арифметические/логические операторы с
приоритетами VBA, ~100 встроенных функций (строки, математика, преобразования,
даты, `Format`, `Split/Join`, `MsgBox/InputBox`, `CreateObject`), константы `vb*`.

### Классы, пользовательские типы, перечисления

- **Модули классов** (`.cls`, заголовок `VERSION 1.0 CLASS` распознаётся сам,
  либо `vba_syntax::parse_class_module`): поля `Public/Private`,
  `Property Get/Let/Set`, `Sub/Function`, `Me`, `Class_Initialize`,
  `New` / `Dim x As New Class`, `TypeOf x Is Class`, `TypeName`.
  Экземпляр — `ClassInstance: Object`; интерпретатор перехватывает вызовы его
  членов (`call_member`), поэтому методы класса получают `ByRef`-аргументы,
  а приватные члены недоступны снаружи (ошибка 438). В отладчике шаги заходят
  в методы класса, в Locals виден `Me` с полями.
- **`Type ... End Type`**: `Value::Udt` со значимой семантикой (копируется при
  присваивании), вложенные типы и массивы в полях, массивы UDT,
  запись `a(i).P.X = 1` и `With a(i) : .X = 1` пишет в саму переменную.
- **`Enum`**: члены — глобальные константы `Long`, доступны и как `Color.Red`;
  переменные `As Color` хранятся как `Long`.
- **`ParamArray`** собирает оставшиеся аргументы в массив `Variant` с базой 0.

## Известные ограничения и план

1. **Классы, дальше**: `Class_Terminate` (нужен счётчик ссылок с доступом к
   интерпретатору), `Implements`, события `WithEvents/RaiseEvent`, член по
   умолчанию (`VB_UserMemId = 0`), `Property` в стандартных модулях.
   `Dim x As New C` создаёт объект сразу, а не при первом обращении.
2. Фиксированные строки `String * N`, `Declare` (Win32 API — через хост).
3. **Модули и видимость**: `Private` пока не изолирует, квалифицированные
   `Module1.Foo`, несколько проектов/ссылок.
4. **Компиляция в байткод** вместо tree-walking: даст `Set Next Statement`
   (Ctrl+F9), точный `Resume Next` для ошибок внутри вложенных блоков,
   отладку без рекурсии Rust-стека и скорость.
5. **Семантика ближе к VBA**: сравнение Variant-строки с числом,
   локаль для дат/чисел, `Format` целиком, `Decimal`, `Erase`, `LSet/RSet`,
   файловый ввод-вывод (`Open/Print #/Line Input`).
6. **Хосты**: эталонный мок Excel (`Workbook/Worksheet/Range` поверх
   `calamine`/`rust_xlsxwriter`), COM-мост на Windows (`IDispatch` через `windows-rs`).
7. **Отладочный протокол**: DAP-сервер поверх `ChannelFrontend` для VS Code.
8. **Совместимость**: корпус реальных макросов + сравнение вывода с настоящим
   Excel (golden-тесты).
