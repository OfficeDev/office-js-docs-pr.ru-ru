---
title: Набор обязательных элементов API для надстройки Outlook 1.1
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 13334029cd30742e6d7dd77cb569a1028a35106a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433035"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Набор обязательных элементов API для надстройки Outlook 1.1

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). 

## <a name="whats-new-in-11"></a>Новые возможности в версии 1.1

Набор обязательных элементов 1.1 включает все возможности набора обязательных элементов версии 1.0. В нем надстройки получили возможность доступа к тексту сообщений и встреч, а также возможность изменения текущего элемента.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [Body](/javascript/api/outlook_1_1/office.body). Предоставляет методы для добавления и изменения содержимого элемента в надстройке Outlook.
- Добавлен объект [Location](/javascript/api/outlook_1_1/office.location). Предоставляет методы, позволяющие получить и задать место проведения собрания в надстройке Outlook.
- Добавлен объект [Recipients](/javascript/api/outlook_1_1/office.recipients). Предоставляет методы, позволяющие получить и задать получателей для встречи или сообщения в надстройке Outlook.
- Добавлен объект [Subject](/javascript/api/outlook_1_1/office.subject). Предоставляет методы, позволяющие получить и задать тему встречи или сообщения в надстройке Outlook.
- Добавлен объект [Time](/javascript/api/outlook_1_1/office.time). Предоставляет методы, позволяющие получить и задать время начала и окончания собрания в надстройке Outlook.
- Добавлен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#addfileattachmentasyncuri-attachmentname-options-callback). Добавляет файл в сообщение или встречу в качестве вложения.
- Добавлен метод [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#additemattachmentasyncitemid-attachmentname-options-callback). Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.
- Добавлен метод [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#removeattachmentasyncattachmentid-options-callback). Удаляет вложение из сообщения или встречи.
- Добавлено свойство [Office.context.mailbox.item.body](office.context.mailbox.item.md#body-bodyjavascriptapioutlook11officebody). Получает объект, предоставляющий методы для работы с текстом элемента.
- Добавлено свойство [Office.context.mailbox.item.bcc](office.context.mailbox.item.md#bcc-recipientsjavascriptapioutlook11officerecipients). Получает или задает получателей скрытой копии сообщения.
- Добавлено свойство [Office.MailboxEnums.RecipientType](/javascript/api/outlook_1_1/office.mailboxenums.recipienttype). Указывает тип получателя для встречи.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)