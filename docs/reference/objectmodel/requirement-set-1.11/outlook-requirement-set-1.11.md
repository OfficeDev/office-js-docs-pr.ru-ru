---
title: Outlook API надстройки 1.11
description: Набор требований 1.11 для Outlook API надстройки.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 56066d7b3a6debaeed365a9ca05a3e894762dea3
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681834"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook API надстройки 1.11

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

## <a name="whats-new-in-111"></a>Что нового в 1.11?

Набор требований 1.11 включает все функции набора [требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). В нем добавлены перечисленные ниже возможности.

- Добавлены новые события [для активации на основе событий.](../../../outlook/autolaunch.md#supported-events)
- Добавлены API SessionData.

### <a name="change-log"></a>Журнал изменений

- Добавлено [Office.context.mailbox.item.sessionData:](office.context.mailbox.item.md#properties)Добавляет новое свойство для управления данными сеанса элемента в режиме Compose.
- Добавлены [Office. SessionData.](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true)Добавляет новый объект, представляю который представляет данные сеанса элемента составить.
- Добавлены новые события [для активации на основе событий.](../../../outlook/autolaunch.md#supported-events)Добавляет поддержку для следующих событий.

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- Добавлены [Office. AppointmentTimeChangedEventArgs:](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)добавляет объект, поддерживаючий `OnAppointmentTimeChanged` событие.
- Добавлены [Office. AttachmentsChangedEventArgs:](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)добавляет объект, поддерживаючий события `OnAppointmentAttachmentsChanged` и `OnMessageAttachmentsChanged` события.
- Добавлены [Office. InfobarClickedEventArgs:](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)Добавляет объект, поддерживаюный `OnInfoBarDismissClicked` событие.
- Добавлены [Office. RecipientsChangedEventArgs:](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)добавляет объект, поддерживаючий события `OnAppointmentAttendeesChanged` и `OnMessageRecipientsChanged` события.
- Добавлены [Office. RecurrenceChangedEventArgs:](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)добавляет объект, поддерживающий `OnAppointmentRecurrenceChanged` событие.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
