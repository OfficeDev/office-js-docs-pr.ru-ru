Outlook надстройки в основном используют API, выставленные через объект [почтовых ящиков.](/javascript/api/outlook/office.mailbox) Чтобы получить объекты и члены специально для использования в надстройках Outlook, такие как объект [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), используйте свойство [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) объекта **Context** для получения доступа к объекту **Mailbox**, как показано в следующей строке кода.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Кроме того, Outlook надстройки могут использовать следующие объекты.

-  Объект **Office** для инициализации.

-  Объект **Context** для получения доступа к контенту и отображения языковых свойств.

-  Объект **RoamingSettings** для сохранения пользовательских свойств, относящихся к надстройке Outlook, в почтовом ящике пользователя, в котором установлено приложение.

Сведения об использовании API Outlook JavaScript см. в Outlook [надстройки.](../outlook/outlook-add-ins-overview.md)