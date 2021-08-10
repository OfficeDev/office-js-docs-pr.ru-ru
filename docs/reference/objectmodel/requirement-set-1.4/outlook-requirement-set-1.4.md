---
title: Набор обязательных элементов API для надстройки Outlook 1.4
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.4.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: fb52a4cb2b1834834f3fde006e4e76f005a044a5e2aefd2cbe789414cea6e29e
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57091893"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Набор обязательных элементов API для надстройки Outlook 1.4

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-14"></a>Новые возможности в версии 1.4

Набор требований 1.4 включает все функции набора [требований 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). В нем добавлен доступ к пространству имен `Office.ui`.

### <a name="change-log"></a>Журнал изменений

- Добавлен [Office.context.ui.displayDialogAsync:](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_)отображает диалоговое окно в Office приложении.
- Добавлен метод [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageParent_message__messageOptions_). Доставляет сообщение из диалогового окна родительской странице.
- Добавлен объект [Dialog](/javascript/api/office/office.dialog). Объект, возвращаемый при вызове метода [`displayDialogAsync`](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_).

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
