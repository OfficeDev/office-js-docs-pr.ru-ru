Надстройки Outlook в основном используют API, предоставляемые через объект [Mailbox](/javascript/api/outlook/Office.mailbox) . Чтобы получить объекты и члены специально для использования в надстройках Outlook, такие как объект [Item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md), используйте свойство [mailbox](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md) объекта **Context** для получения доступа к объекту **Mailbox**, как показано в следующей строке кода.

```js
// Access the Item object.
var item = Office.context.mailbox.item;

```

Кроме того, надстройки Outlook могут использовать следующие объекты:

-  Объект **Office** для инициализации.

-  Объект **Context** для получения доступа к контенту и отображения языковых свойств.

-  Объект **RoamingSettings** для сохранения пользовательских свойств, относящихся к надстройке Outlook, в почтовом ящике пользователя, в котором установлено приложение.

Для получения дополнительных сведений об использовании API JavaScript для Outlook, ознакомьтесь с разделом [надстройки Outlook](../outlook/outlook-add-ins-overview.md).