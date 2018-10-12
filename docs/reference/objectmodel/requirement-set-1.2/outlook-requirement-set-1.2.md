# <a name="outlook-add-in-api-requirement-set-12"></a>Набор требований к API надстройки Outlook 1.2

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора требований](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). 

## <a name="whats-new-in-12"></a>Новые возможности в версии 1.2

Набор требований 1.2 включает все возможности [набора требований версии 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Благодаря ему надстройки теперь могут вставлять текст на месте курсора пользователя (как в теме, так и в тексте сообщения).

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string). Асинхронно возвращает выделенные данные из темы или текста сообщения.
- Добавлен метод [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback). Асинхронно вставляет данные в текст или тему сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata). Добавлено свойство `attachments` параметра `formData`.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata). Добавлено свойство `attachments` параметра `formData`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)