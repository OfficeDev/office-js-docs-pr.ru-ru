---
title: Outlook API надстройки 1.9
description: Набор требований 1.9 для Outlook API надстройки.
ms.date: 05/17/2021
ms.localizationpriority: medium
---

# <a name="outlook-add-in-api-requirement-set-19"></a>Outlook API надстройки 1.9

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-19"></a>Что нового в 1.9?

Набор требований 1.9 включает все функции набора [требований 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для приложений-на-отправке, настраиваемые свойства и функции отображения формы.
- Добавлена поддержка `Dialog.messageChild`.

### <a name="change-log"></a>Журнал изменений

- [Добавлено CustomProperties.getAll](/javascript/api/outlook/office.customproperties?view=outlook-js-1.9&preserve-view=true#outlook-office-customproperties-getall-member(1)): добавляет новую функцию в `CustomProperties` объект, который получает все настраиваемые свойства.
- Добавлен [диалоговое окно.messageChild](../../../develop/dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box): добавляет новый метод, который доставляет сообщение со страницы-организатора, например области задач или файла функций без пользовательского интерфейса, в диалоговое окно, открытое со страницы.
- Добавлен [элемент манифеста ExtendedPermissions](../../manifest/extendedpermissions.md): добавляет детский элемент в элемент [манифеста VersionOverrides](../../manifest/versionoverrides.md) . Чтобы надстройка поддержала функцию [приложения-на-отправке](../../../outlook/append-on-send.md), `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.
- [Добавлена Office.context.mailbox.displayAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displayappointmentformasync-member(1)): `Mailbox` добавляет новую функцию к объекту, который отображает существующую встречу. Это версия async метода `displayAppointmentForm` .
- Добавлен [Office.context.mailbox.displayMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaymessageformasync-member(1)): `Mailbox` добавляет новую функцию к объекту, который отображает существующее сообщение. Это версия async метода `displayMessageForm` .
- Добавлен [Office.context.mailbox.displayNewAppointmentFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewappointmentformasync-member(1)): `Mailbox` добавляет новую функцию к объекту, который отображает новую форму встречи. Это версия async метода `displayNewAppointmentForm` .
- Добавлен [Office.context.mailbox.displayNewMessageFormAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true#outlook-office-mailbox-displaynewmessageformasync-member(1)): `Mailbox` добавляет новую функцию к объекту, который отображает новую форму сообщения. Это версия async метода `displayNewMessageForm` .
- Добавлен [Office.context.mailbox.item.body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-1.9&preserve-view=true#outlook-office-body-appendonsendasync-member(1)): `Body` добавляет новую функцию к объекту, который добавляет данные в конец тела элемента в режиме Compose.
- Добавлен [Office.context.mailbox.item.displayReplyAllFormAsync](office.context.mailbox.item.md#methods): `Item` добавляет новую функцию к объекту, который отображает форму "Ответить все" в режиме Чтения. Это версия async метода `displayReplyAllForm` .
- Добавлен [Office.context.mailbox.item.displayReplyFormAsync](office.context.mailbox.item.md#methods): `Item` добавляет новую функцию к объекту, отображаемой в режиме "Ответ". Это версия async метода `displayReplyForm` .

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
