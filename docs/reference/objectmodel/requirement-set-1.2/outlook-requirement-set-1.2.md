---
title: Набор обязательных элементов API для надстройки Outlook 1.2
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: a77c9c0399d42029b86595f119c51071dc3bf4ac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433546"
---
# <a name="outlook-add-in-api-requirement-set-12"></a>Набор обязательных элементов API для надстройки Outlook 1.2

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets). 

## <a name="whats-new-in-12"></a>Новые возможности в версии 1.2

Набор обязательных элементов 1.2 включает все возможности [набора обязательных элементов версии 1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md). Благодаря ему надстройки теперь могут вставлять текст на месте пользовательского указателя (как в теме, так и в тексте сообщения).

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.mailbox.item.getSelectedDataAsync](office.context.mailbox.item.md#getselecteddataasynccoerciontype-options-callback--string). Асинхронно возвращает данные, выбранные в теме или тексте сообщения.
- Добавлен метод [Office.context.mailbox.item.setSelectedDataAsync](office.context.mailbox.item.md#setselecteddataasyncdata-options-callback). Асинхронно вставляет данные в текст или тему сообщения.
- Изменен метод [Office.context.mailbox.item.displayReplyAllForm](office.context.mailbox.item.md#displayreplyallformformdata). Добавлено свойство `attachments` параметра `formData`.
- Изменен метод [Office.context.mailbox.item.displayReplyForm](office.context.mailbox.item.md#displayreplyformformdata). Добавлено свойство `attachments` параметра `formData`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](https://docs.microsoft.com/outlook/add-ins/quick-start)