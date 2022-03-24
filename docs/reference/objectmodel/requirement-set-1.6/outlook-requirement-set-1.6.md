---
title: Набор обязательных элементов API для надстройки Outlook 1.6
description: Функции и API, которые были Outlook надстройки и Office API JavaScript в рамках API почтовых ящиков 1.6.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: c9bb9194b300f95f105e24f5e318e348055040ea
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746111"
---
# <a name="outlook-add-in-api-requirement-set-16"></a>Набор обязательных элементов API для надстройки Outlook 1.6

Подмножество API Outlook надстройки в API Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-16"></a>Новые возможности в версии 1.6

Набор требований 1.6 включает все функции набора [требований 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для контекстных надстроек, которые позволяют получить соответствие объекта или RegEx, выбранного пользователем для активации надстройки.
- Добавлен новый интерфейс API для открытия новой формы сообщения.
- Добавлена возможность надстройки определять тип учетной записи почтового ящика пользователя.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#methods). Добавляет новую функцию, которая возвращает объекты, найденные в выделенном совпадении. Выделенные совпадения применяются к контекстным надстройкам.
- Добавлен объект [Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#methods). Добавляет новую функцию, которая возвращает строковые значения в выделенном совпадении, соответствующие регулярным выражениям, определенным в XML-файле манифеста. Выделенные совпадения применяются к контекстным надстройкам.
- Добавлен объект [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#methods). Добавляет новую функцию, которая открывает новую форму сообщения.
- Добавлен объект [Office.context.mailbox.userProfile.accountType](/javascript/api/outlook/office.userprofile?view=outlook-js-1.6&preserve-view=true#outlook-office-userprofile-accounttype-member). Добавляет новый элемент в профиль пользователя, указывающий тип учетной записи пользователя.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
