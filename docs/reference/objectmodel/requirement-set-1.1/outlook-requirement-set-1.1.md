---
title: Набор требований к API надстройки Outlook 1.1
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 2ecd337cd838cd6dd9deb4fe5e77ee789106f3f9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165470"
---
# <a name="outlook-add-in-api-requirement-set-11"></a>Набор обязательных элементов API для надстройки Outlook 1.1

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook. API JavaScript для Outlook 1,1 (почтовый ящик 1,1) — это первая версия API.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-11"></a>Новые возможности в версии 1.1

Набор требований 1,1 включает все [Общие наборы требований API](../../requirement-sets/office-add-in-requirement-sets.md) , поддерживаемые в Outlook. В нем надстройки получили возможность доступа к тексту сообщений и встреч, а также возможность изменения текущего элемента.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [Body](/javascript/api/outlook/office.body?view=outlook-js-1.1). Предоставляет методы для добавления и изменения содержимого элемента в надстройке Outlook.
- Добавлен объект [Location](/javascript/api/outlook/office.location?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать место проведения собрания в надстройке Outlook.
- Добавлен объект [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать получателей для встречи или сообщения в надстройке Outlook.
- Добавлен объект [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать тему встречи или сообщения в надстройке Outlook.
- Добавлен объект [Time](/javascript/api/outlook/office.time?view=outlook-js-1.1). Предоставляет методы, позволяющие получить и задать время начала и окончания собрания в надстройке Outlook.
- Добавлен метод [Office.context.mailbox.item.addFileAttachmentAsync](office.context.mailbox.item.md#methods). Добавляет файл в сообщение или встречу в качестве вложения.
- Добавлен метод [Office.context.mailbox.item.addItemAttachmentAsync](office.context.mailbox.item.md#methods). Добавляет к сообщению или встрече элемент Exchange (например, сообщение) в виде вложения.
- Добавлен метод [Office.context.mailbox.item.removeAttachmentAsync](office.context.mailbox.item.md#methods). Удаляет вложение из сообщения или встречи.
- Добавлено свойство [Office.context.mailbox.item.body](office.context.mailbox.item.md#properties). Получает объект, предоставляющий методы для работы с текстом элемента.
- Добавлена строка [Office. Context. Mailbox. Item. BCC](office.context.mailbox.item.md#properties) сообщения.
- Добавлено свойство [Office.MailboxEnums.RecipientType](/javascript/api/outlook/office.mailboxenums.recipienttype?view=outlook-js-1.1). Указывает тип получателя для встречи.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
