# <a name="outlook-add-in-api-requirement-set-16"></a>Набор обязательных элементов API для надстроек Outlook 1.6

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-16"></a>Новые возможности в версии 1.6

Набор обязательных элементов включает все возможности [набора обязательных элементов 1.6](../requirement-set-1.5/outlook-requirement-set-1.5.md). Добавлены следующие возможности:

- добавлены новые API для контекстных надстроек, позволяющие получить соответствие сущности или RegEx, выбранной пользователем для активации надстройки.
- Добавлен новый интерфейс API для открытия новой формы сообщения.
- Добавлена возможность надстройки определить тип учетной записи почтового ящика пользователя.

### <a name="change-log"></a>Журнал изменений

- [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities) — добавляет новую функцию, которая возвращает объекты, найденные в выделенном совпадении. Выделенные совпадения применяются к контекстным надстройкам.
- [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object) — добавляет новую функцию, которая возвращает значения строки в выделенном совпадении, которые соответствуют регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к контекстным надстройкам.
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters) — добавляет новую функцию, которая открывает новую форму сообщения.
- [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string) — добавляет новый член в профиль пользователя, который указывает тип учетной записи пользователя.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)