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
| `vba-excel` | Эталонный хост: объектная модель Excel в памяти поверх `Host`/`Object`, чтение и запись `.xlsx`. |
| `vba-cli` | `vba run Main.bas Class1.cls ...`, `vba debug ...` — консольный отладчик с командами в духе VBE, брейкпоинты вида `Module:line`. |

## Система типов (`value.rs`)

`Value` = VBA Variant: `Empty, Null, Boolean, Byte, Integer(i16), Long(i32),
LongLong(i64), Single, Double, Currency(i64 ×10⁴), Decimal (rust_decimal),
Date(f64 OLE), String, Object(Option<ObjectRef>) (None = Nothing), Array, Error, Missing`.

- Объявленный тип (`Dim x As Long`) хранится в слоте переменной; присваивание
  приводит значение через `VarType::coerce` (банковское округление, Overflow = ошибка 6).
- Интерпретатор выводит статический тип операндов (литералы с суффиксами
  `% & ^ ! # @`, типизированные переменные, `CInt()` и т. п.). Если оба операнда
  типизированы, переполнение — ошибка 6 (`Dim i As Integer: i = 32767: i + 1`);
  если хоть один Variant, тип расширяется (Byte → Integer → Long → Double).
- `+` на двух строках — конкатенация, `/` — Double (Single для Single),
  `\` и `Mod` округляют операнды до Long по-банковски.
- `Null` распространяется через операторы; `And/Or/Imp` с Null — трёхзначная
  логика (`False And Null` = False). Число Variant всегда меньше строки Variant.
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
- **`Class_Terminate`**: `ClassInstance` реализует `Drop`. Когда пропадает
  последняя ссылка, его поля переезжают в «заместителя» `Me`, а тот попадает в
  очередь. Интерпретатор разбирает очередь после каждого оператора и после
  выхода из процедуры, поэтому `Set x = Nothing` или `End Sub` запускают
  `Class_Terminate` сразу, и отладчик в него заходит. При `End` и Reset
  `Class_Terminate` не вызывается, как в VBA. Циклические ссылки, как и в VBA,
  не освобождаются.
- **`Implements`**: `shape.Area` у класса без собственного `Area` вызывает
  `IShape_Area`. Тип переменной (`Dim s As IShape`) не хранится, поэтому
  собственный публичный член с тем же именем побеждает интерфейсный.
  `TypeOf x Is IShape` учитывает интерфейсы.
- **События**: `Public Event Changed(...)`, `RaiseEvent Changed(x)`,
  `Private WithEvents acc As Account` и обработчик `acc_Changed`.
  `Set acc = a` подписывает объект слабой ссылкой. Аргументы передаются ByRef,
  поэтому работает `Cancel = True`. После `Set acc = Nothing` или переназначения
  обработчик больше не вызывается.
- **Член по умолчанию** (`Attribute Item.VB_UserMemId = 0`): `obj(2)`,
  `x = obj + 1`, `Debug.Print obj`, `If obj Then`, `CStr(obj)`, `obj = 5`.
  **`For Each`** по своей коллекции: член с `VB_UserMemId = -4`, обычно
  `Set NewEnum = mItems.[_NewEnum]`. Идентификаторы в `[скобках]` поддерживаются.
- **`Dim x As New C`** создаёт объект при первом обращении и заново после
  `Set x = Nothing`, поэтому `x Is Nothing` всегда `False`, как в VBA.
- **Стандартные модули**: `Property Get/Let/Set` в `.bas`; обращения
  `Module1.Foo`, `Module1.Counter = 5`. `Private` процедуры и переменные модуля
  видны только внутри него, а одноимённые приватные процедуры в разных модулях
  не конфликтуют. Хост через `Engine::run` может вызвать и приватный макрос.

### Файлы и строковые операторы

- **Файловый ввод-вывод**: `Open ... For Input/Output/Append As #n`,
  `Print #` (с теми же `;`, `,`, `Tab(n)`, `Spc(n)`, что у `Debug.Print`),
  `Write #` (строки в кавычках, `#TRUE#`, `#2024-03-01#`), `Line Input #`,
  `Input #` (разбирает то, что пишет `Write #`), `Close`, `EOF`, `LOF`,
  `FreeFile`, а также `Dir` с шаблонами и `Dir()` для продолжения, `Kill`,
  `FileLen`, `FileCopy`, `MkDir`. Строки пишутся в UTF-8 с концами `\r\n`.
- **Доступ к файлам идёт через `Host`** (`read_file`, `write_file`,
  `delete_file`, `create_dir`, `list_dir`) и по умолчанию запрещён
  (ошибка 75), чтобы встроенный движок не трогал диск без разрешения.
  `StdHost` и `ExcelHost` работают с настоящей файловой системой через
  `host::fs`, у `CaptureHost` она в памяти (`files`, `file_text`) для тестов.
  Файл для чтения читается целиком при `Open`, вывод копится и пишется при
  `Close` или в конце макроса. Режимы `Binary`/`Random` (`Get`/`Put`) пока не
  поддерживаются.
- **`String * N`** — `VarType::FixedString(N)`: присваивание обрезает или
  дополняет пробелами, работает и в полях `Type`.
- **`Erase`** очищает элементы массива фиксированного размера и освобождает
  динамический массив. **`LSet`/`RSet`** выравнивают значение в текущей длине
  строки. **`Mid(s, start, len) = v`** заменяет символы на месте.

### Мок Excel (`vba-excel`)

`ExcelHost` реализует `Host`: глобальные `Application`, `ActiveWorkbook`,
`ThisWorkbook`, `ActiveSheet`, `ActiveCell`, `Workbooks`, `Worksheets`,
`Sheets`, `WorksheetFunction` и ~60 констант `xl*`; `global_object` —
`Application`, поэтому неквалифицированные `Range("A1")`, `Cells(i, 1)`,
`Rows.Count` работают как в Excel.

- **Данные** (`model.rs`): лист — разреженный `BTreeMap<(строка, столбец), CellData>`
  (значение, формула, стиль), числа хранятся как `Double`, как в Excel.
  `End` повторяет Ctrl+стрелка, `CurrentRegion` и `UsedRange` считаются по
  заполненным ячейкам. Объекты (`RangeObj`, `WorksheetObj`, ...) — лёгкие
  ссылки на общие данные, поэтому `Set r = Range("A1")` видит последующие изменения.
- **Range**: `Value`/`Value2` (скаляр или массив `(1 To n, 1 To m)`, запись
  скаляра, 1D и 2D массивов), `Formula` (хранится и сохраняется, не вычисляется),
  `Text` через `NumberFormat`, `Address(RowAbsolute, ColumnAbsolute)`, `Cells`,
  `Range`, `Rows`, `Columns`, `EntireRow/Column`, `Offset`, `Resize`, `End`,
  `CurrentRegion`, `Count`, `ClearContents`, `Clear`, `Copy`, `Font.Bold/Italic/Size/Color`,
  `Interior.Color`, `For Each` по ячейкам, строкам и столбцам.
- **Листы и книги**: `Worksheets.Add(Before/After/Count)`, `Name`, `Index`,
  `Delete`, `Activate`, `Workbooks.Add/Open`, `SaveAs`, `Save`, `Close`.
- **WorksheetFunction**: `Sum`, `Product`, `Average`, `Min`, `Max`, `Large`,
  `Small`, `Count`, `CountA`, `CountBlank`, `CountIf`, `SumIf`, `AverageIf`
  (критерии `">5"`, `"<>x"`, подстановочные `*`/`?`), `VLookup`, `HLookup`,
  `Match`, `Index`, `Round/RoundUp/RoundDown`, `Trim`, `Transpose`. Ошибка —
  1004 «Unable to get the X property of the WorksheetFunction class», как в Excel.
- **`.xlsx`** (фича `xlsx`, по умолчанию включена): чтение через `calamine`
  (значения, формулы, даты), запись через `rust_xlsxwriter` (значения, формулы,
  жирный/курсив, цвет и размер шрифта, заливка, числовой формат).
- **CLI**: `--excel` (пустая книга), `--workbook in.xlsx`, `--save out.xlsx`;
  отладчик работает так же, `? Range("A1").Value` в Immediate.

Ограничения мока: формулы не пересчитываются; стили при чтении `.xlsx` не
загружаются; формат целых строк/столбцов применяется только к существующим
ячейкам; `Union` несмежных диапазонов приближается охватывающим прямоугольником;
нет `Names`, `ListObjects`, `Charts`, событий листа.

### Встроенные функции (`builtins/`)

Разбиты по библиотекам VBA: `strings.rs`, `conversion.rs` (+ Information),
`math.rs` (+ финансовые `Pmt`, `IRR`, `NPV`…), `datetime.rs`, `format.rs`
(`Format` целиком: секции, литералы, именованные форматы, маски дат,
`FormatNumber/Currency/Percent/DateTime`). Формы с `$` (`Left$`) возвращают
String и дают ошибку 94 на Null. `Rnd`/`Randomize` — настоящий 24-битный
генератор VBA, последовательности совпадают с Office.

### Тесты соответствия (`tests/vba/*.bas`)

Около 2000 проверок, написанных на самом VBA: `Check`, `CheckTyped`,
`CheckErr` из `Assert.bas`. Ожидаемые значения взяты из примеров справочника
VBA от Microsoft (MicrosoftDocs/VBA-Docs, код под MIT) и из известного
поведения VBA 7 (en-US). `tests/conformance.rs` запускает каждый файл отдельным
тестом. Чтобы добавить проверку, достаточно дописать строку в нужный `.bas`.

### `Debug.Print`

Печать как в VBA: числа с пробелом под знак впереди и пробелом после,
`;` склеивает, `,` переходит к следующей зоне в 14 символов, завершающий
`;` или `,` оставляет строку открытой для следующего `Debug.Print`.

## Известные ограничения и план

1. **Классы, дальше**: события от объектов хоста (`WithEvents` на объекте
   Excel), диспетчеризация `Implements` по объявленному типу переменной.
2. `Declare` (Win32 API — через хост), файлы в режимах `Binary`/`Random`.
3. **Модули и видимость**: несколько проектов/ссылок, `Option Private Module`.
4. **Компиляция в байткод** вместо tree-walking: даст `Set Next Statement`
   (Ctrl+F9), точный `Resume Next` для ошибок внутри вложенных блоков,
   отладку без рекурсии Rust-стека и скорость.
5. **Семантика ближе к VBA**: локаль для дат/чисел (сейчас en-US).
6. **Хосты**: мок Excel дальше (вычисление формул, `Find`, `Sort`,
   `AutoFilter`, `Names`), COM-мост на Windows (`IDispatch` через `windows-rs`)
   для запуска тех же макросов в настоящем Excel.
7. **Отладочный протокол**: DAP-сервер поверх `ChannelFrontend` для VS Code.
8. **Совместимость**: корпус реальных макросов + сравнение вывода с настоящим
   Excel (golden-тесты).
