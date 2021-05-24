---
title: Набор требований к API надстройки Outlook 1.2
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.2.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: d643f0fdf07c5f22d8d863075b894cfc05b21363
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590402"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Набор требований к API надстройки Outlook 1.2

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-12"></a>Новые возможности в версии 1.2

Набор требований 1.2 включает все функции набора [требований 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Благодаря ему надстройки теперь могут вставлять текст на месте пользовательского указателя (как в теме, так и в тексте сообщения).

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#methods). Асинхронно возвращает данные, выбранные в теме или тексте сообщения.
- Добавлен метод [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#methods). Асинхронно вставляет данные в текст или тему сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#methods). Добавлено свойство `attachments` параметра `formData`.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#methods). Добавлено свойство `attachments` параметра `formData`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
