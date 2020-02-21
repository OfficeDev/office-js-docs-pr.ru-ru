---
title: Набор обязательных элементов API для надстройки Outlook 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: cb4c8eecd63604aa633ade1a40eb5391b3a62ef2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165407"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Набор обязательных элементов API для надстройки Outlook 1.4

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-14"></a>Новые возможности в версии 1.4

Набор обязательных элементов 1.4 включает все возможности [набора обязательных элементов версии 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). В нем добавлен доступ к пространству имен `Office.ui`.

### <a name="change-log"></a>Журнал изменений

- Добавлен метод [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-). Отображает диалоговое окно в ведущем приложении Office.
- Добавлен метод [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-). Доставляет сообщение из диалогового окна родительской странице.
- Добавлен объект [Dialog](/javascript/api/office/office.dialog). Объект, возвращаемый при вызове метода [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-).

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
