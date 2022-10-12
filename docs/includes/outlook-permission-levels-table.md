|**Каноническое</br>имя уровня разрешений**|**Имя XML-манифеста**|**Имя манифеста Teams**|**Краткое описание**|
|:-----|:-----|:-----|:-----|
|**Ограничен**|Restricted|MailboxItem.Restricted.User|Разрешает использование сущностей, но не регулярных выражений. |
|**чтение элемента**|ReadItem|MailboxItem.Read.User|Помимо того, что разрешено в **ограниченном доступе**, он позволяет:<ul><li>регулярные выражения;</li><li>доступ на чтение API надстроек Outlook;</li><li>получение свойств элемента и маркера обратного вызова.</li></ul> |
|**Чтение и запись элемента**|ReadWriteItem|MailboxItem.ReadWrite.User|Помимо того, что разрешено в **элементе чтения**, он позволяет:<ul><li>полный доступ ко всем элементам API Outlook, кроме метода `makeEwsRequestAsync`;</li><li>задание свойств элемента.</li></ul> |
|**Чтение и запись почтового ящика**|ReadWriteMailbox|Mailbox.ReadWrite.User|Помимо того, что разрешено в элементе чтения **и записи**, он позволяет:<ul><li>создание, чтение и запись элементов и папок;</li><li>отправка папок;</li><li>вызов метода [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).</li></ul> |

Разрешения объявляются в манифесте. Разметка зависит от типа манифеста.

- **XML-манифест**: используйте **\<Permissions\>** элемент.
- **Манифест Teams (предварительная версия)**: используйте свойство name объекта в массиве authorization.permissions.resourceSpecific.

> [!NOTE]
>
> - Для надстроек, использующих функцию добавления при отправке, требуется дополнительное разрешение. С помощью XML-манифеста вы указываете разрешение в [элементе ExtendedPermissions](/javascript/api/manifest/extendedpermissions) . Дополнительные сведения см. в статье ["Реализация добавления при отправке" в надстройке Outlook](../outlook/append-on-send.md). В манифесте Teams (предварительная версия) это разрешение указывается с именем **Mailbox.AppendOnSend.User** в дополнительном объекте в массиве authorization.permissions.resourceSpecific.
> - Для надстроек, использующих общие папки, требуется дополнительное разрешение. С помощью XML-манифеста вы указываете разрешение, задав для элемента [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) значение `true`. Дополнительные сведения см. в статье "Включение общих папок и сценариев общих [почтовых ящиков в надстройке Outlook"](../outlook/delegate-access.md). В манифесте Teams (предварительная версия) это разрешение указывается с именем **Mailbox.SharedFolder** в дополнительном объекте в массиве authorization.permissions.resourceSpecific.
