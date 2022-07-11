Надстройки Outlook, в основном, используют набор API, предоставляемый через объект [Mailbox](/javascript/api/outlook/office.mailbox). Чтобы получить доступ к объектам и членам, предназначенным специально для использования в надстройке Outlook, например [объекте Item](/javascript/api/outlook/office.item), [](/javascript/api/office/office.context#office-office-context-mailbox-member) используйте свойство почтового ящика объекта **Context** для доступа к объекту **почтового** ящика, как показано в следующей строке кода.

```js
// Access the Item object.
const item = Office.context.mailbox.item;
```

Кроме того, надстройки Outlook могут использовать следующие объекты.

- Объект **Office** для инициализации.

- Объект **Context** для получения доступа к контенту и отображения языковых свойств.

- Объект **RoamingSettings** для сохранения пользовательских свойств, относящихся к надстройке Outlook, в почтовом ящике пользователя, в котором установлено приложение.

Сведения об использовании JavaScript в надстройках Outlook см. в статье [Надстройки Outlook](../outlook/outlook-add-ins-overview.md).
