
# <a name="diagnostics"></a>diagnostics

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics

Предоставляет надстройке Outlook диагностические сведения.

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="members-and-methods"></a>Члены и методы

| Член | Тип |
|--------|------|
| [hostName](#hostname-string) | Член |
| [hostVersion](#hostversion-string) | Член |
| [OWAView](#owaview-string) | Член |

### <a name="members"></a>Члены

####  <a name="hostname-string"></a>hostName :String

Получает строку, представляющую имя ведущего приложения.

Строка, которая может быть одним из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.

##### <a name="type"></a>Тип:

*   Строка

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

####  <a name="hostversion-string"></a>hostVersion :String

Получает строку, которая представляет версию ведущего приложения или Exchange Server.

Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения — Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Например, строка `15.0.468.0`.

##### <a name="type"></a>Тип:

*   String

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

####  <a name="owaview-string"></a>OWAView :String

Получает строку, отображающую текущее представление Outlook Web App.

Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.

Если ведущее приложение — не Outlook Web App, тогда при получении доступа к этому свойству будет выдаваться значение `undefined`.

Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов:

*   `OneColumn`, которое используется в случае узкого экрана. Outlook Web App использует этот макет размером в один столбец на экране смартфона.
*   `TwoColumns`, которое используется в случае более широкого экрана. Outlook Web App использует это представление на большинстве планшетных ПК.
*   `ThreeColumns`, которое используется в случае широкого экрана. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.

##### <a name="type"></a>Тип:

*   Строка

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора обязательных элементов для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|