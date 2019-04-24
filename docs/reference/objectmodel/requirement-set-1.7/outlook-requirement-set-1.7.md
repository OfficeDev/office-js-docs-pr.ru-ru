---
title: Набор обязательных элементов API для надстройки Outlook 1.7
description: ''
ms.date: 03/20/2019
localization_priority: Priority
ms.openlocfilehash: 8daf10239a704206d53a544185e030afa6b6a27a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450284"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Набор обязательных элементов API для надстройки Outlook 1.7

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

## <a name="whats-new-in-17"></a>Новые возможности в версии 1.7

Набор обязательных элементов 1.7 включает все возможности [набора обязательных элементов версии 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для расписания повторения встреч и сообщений с приглашением на собрание.
- Изменено свойство item.from — теперь оно также доступно в режиме создания.
- Добавлена поддержка событий RecurrenceChanged, RecipientsChanged и AppointmentTimeChanged.

### <a name="change-log"></a>Журнал изменений

- Добавлен объект [From](/javascript/api/outlook_1_7/office.from). Добавляет новый объект, предоставляющий метод получения значения отправителя.
- Добавлен объект [Organizer](/javascript/api/outlook_1_7/office.organizer). Добавляет новый объект, предоставляющий метод получения значения организатора.
- Добавлен объект [Recurrence](/javascript/api/outlook_1_7/office.recurrence). Добавляет новый объект, предоставляющий методы получения и установки расписания повторения встреч и методы получения расписания повторения сообщений с приглашением на собрание.
- Добавлен объект [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone). Добавляет новый объект, представляющий настройку часового пояса расписания повторения.
- Добавлен объект [SeriesTime](/javascript/api/outlook_1_7/office.seriestime). Добавляет новый объект, предоставляющий методы получения и установки даты и времени встреч в повторяющемся ряду и методы получения даты и времени приглашений на собрание в повторяющемся ряду.
- Добавлен объект [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback). Добавляет новый метод, добавляющий обработчик для поддерживаемого события.
- Изменен объект [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsfrom). Добавляет возможность получения значения отправителя в режиме создания.
- Изменен объект [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsorganizer). Добавляет возможность получения значения организатора в режиме создания.
- Добавлен объект [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrence). Добавляет новое свойство, которое получает или задает объект, предоставляющий методы управления расписанием повторения встреч. Это свойство можно также использовать для получения расписания повторения приглашения на собрание.
- Добавлен объект [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-options-callback). Добавляет новый метод, удаляющий обработчиков событий для поддерживаемого типа события. 
- Добавлен объект [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string). Добавляет новое свойство, получающее идентификатор ряда, к которому относится событие.
- Добавлен объект [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days). Добавляет новое перечисление, указывающее день недели или тип дня.
- Добавлен объект [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month). Добавляет новое перечисление, указывающее месяц.
- Добавлен объект [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone). Добавляет новое перечисление, указывающее часовой пояс повторения.
- Добавлен объект [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype). Добавляет новое перечисление, указывающее тип повторения.
- Добавлен объект [ Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber). Добавляет новое перечисление, указывающее неделю месяца.
- Изменен объект [Office.EventType](/javascript/api/office/office.eventtype). Добавляет поддержку событий `RecurrenceChanged`, `RecipientsChanged` и `AppointmentTimeChanged`.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](/outlook/add-ins/)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](/outlook/add-ins/quick-start)
