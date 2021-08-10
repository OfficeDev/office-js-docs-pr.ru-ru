---
title: Набор обязательных элементов API для надстройки Outlook 1.7
description: Обзор API Outlook надстройки (набор требований 1.7)
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0c9a008718b6ef4af0ad7ab1fb7fa059bf16e127e33b306d7a4a4d2a3ba8de8a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089670"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Набор обязательных элементов API для надстройки Outlook 1.7

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!NOTE]
> В этой документации рассматривается не последняя версия [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).

## <a name="whats-new-in-17"></a>Новые возможности в версии 1.7

Набор требований 1.7 включает все функции набора [требований 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для расписания повторения встреч и сообщений с приглашением на собрание.
- Изменено свойство item.from — теперь оно также доступно в режиме создания.
- Добавлена поддержка событий RecurrenceChanged, RecipientsChanged и AppointmentTimeChanged.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [From](/javascript/api/outlook/office.from?view=outlook-js-1.7&preserve-view=true). Добавляет новый объект, предоставляющий метод получения значения отправителя.
- Добавлен объект [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7&preserve-view=true). Добавляет новый объект, предоставляющий метод получения значения организатора.
- Добавлен объект [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7&preserve-view=true). Добавляет новый объект, предоставляющий методы получения и установки расписания повторения встреч и методы получения расписания повторения сообщений с приглашением на собрание.
- Добавлен объект [RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone?view=outlook-js-1.7&preserve-view=true). Добавляет новый объект, представляющий настройку часового пояса расписания повторения.
- Добавлен объект [SeriesTime](/javascript/api/outlook/office.seriestime?view=outlook-js-1.7&preserve-view=true). Добавляет новый объект, предоставляющий методы получения и установки даты и времени встреч в повторяющемся ряду и методы получения даты и времени приглашений на собрание в повторяющемся ряду.
- Добавлен объект [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, добавляющий обработчик для поддерживаемого события.
- Изменен объект [Office.context.mailbox.item.from](office.context.mailbox.item.md#properties). Добавляет возможность получения значения отправителя в режиме создания.
- Изменен объект [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#properties). Добавляет возможность получения значения организатора в режиме создания.
- Добавлен объект [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#properties). Добавляет новое свойство, которое получает или задает объект, предоставляющий методы управления расписанием повторения встреч. Это свойство можно также использовать для получения расписания повторения приглашения на собрание.
- Добавлен объект [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#methods). Добавляет новый метод, удаляющий обработчиков событий для поддерживаемого типа события. 
- Добавлен объект [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#properties). Добавляет новое свойство, получающее идентификатор ряда, к которому относится событие.
- Добавлен объект [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days?view=outlook-js-1.7&preserve-view=true). Добавляет новое перечисление, указывающее день недели или тип дня.
- Добавлен объект [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month?view=outlook-js-1.7&preserve-view=true). Добавляет новое перечисление, указывающее месяц.
- Добавлен объект [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone?view=outlook-js-1.7&preserve-view=true). Добавляет новое перечисление, указывающее часовой пояс повторения.
- Добавлен объект [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype?view=outlook-js-1.7&preserve-view=true). Добавляет новое перечисление, указывающее тип повторения.
- Добавлен объект [ Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber?view=outlook-js-1.7&preserve-view=true). Добавляет новое перечисление, указывающее неделю месяца.
- Изменен объект [Office.EventType](/javascript/api/office/office.eventtype). Добавляет поддержку событий `RecurrenceChanged`, `RecipientsChanged` и `AppointmentTimeChanged`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
