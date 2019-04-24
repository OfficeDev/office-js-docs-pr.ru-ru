---
title: Набор обязательных элементов API для надстройки Outlook 1.6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 0e1f920c259ca1ef8a137bab07132b015d9c75d2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451733"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Набор обязательных элементов API для надстройки Outlook 1.6

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets).

## <a name="whats-new-in-16"></a>Новые возможности в версии 1.6

Набор обязательных элементов 1.6 включает все возможности [набора обязательных элементов версии 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для контекстных надстроек, которые позволяют получить соответствие объекта или RegEx, выбранного пользователем для активации надстройки.
- Добавлен новый интерфейс API для открытия новой формы сообщения.
- Добавлена возможность надстройки определять тип учетной записи почтового ящика пользователя.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entities). Добавляет новую функцию, которая возвращает объекты, найденные в выделенном совпадении. Выделенные совпадения применяются к контекстным надстройкам.
- Добавлен объект [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object). Добавляет новую функцию, которая возвращает строковые значения в выделенном совпадении, соответствующие регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к контекстным надстройкам.
- Добавлен объект [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters). Добавляет новую функцию, которая открывает новую форму сообщения.
- Добавлен объект [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string). Добавляет новый элемент в профиль пользователя, указывающий тип учетной записи пользователя.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/outlook/add-ins/quick-start)
