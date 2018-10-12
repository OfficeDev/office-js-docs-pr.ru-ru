# <a name="outlook-add-in-api-requirement-set-15"></a>Набор требований API для надстроек Outlook 1.5

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-15"></a>Новые возможности в версии 1.5

Набор требований 1.5 включает все возможности [набора требований версии 1.4](../requirement-set-1.4/outlook-requirement-set-1.4.md). К нему добавлены перечисленные ниже возможности.

- Добавлена поддержка [закрепляемых областей задач](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane).
- Добавлена поддержка вызовов [интерфейсов REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api).
- Добавлена возможность отметить вложение как встроенное.
- Добавлена возможность закрыть область задач или диалоговое окно.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.addHandlerAsync](office.context.mailbox.md#addhandlerasynceventtype-handler-options-callback). Добавляет обработчик для поддерживаемого события.
- Добавлен [Office.EventType](office.md#eventtype-string). Указывает событие, связанное с обработчиком событий, и включает поддержку события ItemChanged.
- Добавлен метод [Office.context.mailbox.restUrl](office.context.mailbox.md#resturl-string). Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.
- Изменен метод [Office.context.mailbox.getCallbackTokenAsync](office.context.mailbox.md#getcallbacktokenasyncoptions-callback). Добавлен новый вариант этого метода с новой подписью (`getCallbackTokenAsync([options], callback)`). Исходная версия по-прежнему доступна и осталась без изменений.
- Добавлен метод [Office.context.ui.closeContainer](/javascript/api/office/office.ui#closecontainer--).
- Изменен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback). Новое значение в словаре `options` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata). Новое значение в словаре `formData.attachments` — `isInline`. Оно указывает на то, что изображение встроено в текст сообщения.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)