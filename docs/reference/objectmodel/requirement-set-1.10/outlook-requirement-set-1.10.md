---
title: Outlook API надстройки 1.10
description: Набор требований 1.10 для Outlook API надстройки.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 9e3e30590279036a08a93d8643cd56c2c73be78c
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938409"
---
# <a name="outlook-add-in-api-requirement-set-110"></a>Outlook API надстройки 1.10

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

## <a name="whats-new-in-110"></a>Что нового в 1.10?

Набор требований 1.10 включает все функции набора [требований 1.9.](../requirement-set-1.9/outlook-requirement-set-1.9.md) В нем добавлены перечисленные ниже возможности.

- Добавлены новые API для [активации на](../../../outlook/autolaunch.md) основе событий и функций подписи почты.
- Добавлена возможность включить настраиваемые действия в сообщение уведомления.

### <a name="change-log"></a>Журнал изменений

- Добавлена [точка расширения LaunchEvent.](../../manifest/extensionpoint.md#launchevent)Добавлен новый поддерживаемый тип ExtensionPoint. Он настраивает функции активации на основе событий.
- Добавлен [элемент манифеста LaunchEvents:](../../manifest/launchevents.md)добавляет элемент манифеста для поддержки настройки функции активации на основе событий.
- Измененный [элемент манифеста runtimes:](../../manifest/runtimes.md)добавляет Outlook поддержку. Он ссылается на ФАЙЛЫ HTML и JavaScript, необходимые для функций активации на основе событий.
- Добавлен [Office.context.mailbox.item.body.setSignatureAsync:](/javascript/api/outlook/office.body?view=outlook-js-1.10&preserve-view=true#setSignatureAsync_data__options__callback_)добавляет новую функцию в `Body` объект. Он добавляет или заменяет подпись в корпусе элемента в режиме Compose.
- Добавлена [Office.context.mailbox.item.disableClientSignatureAsync:](office.context.mailbox.item.md#methods)добавляется новая функция, которая отключает подпись клиента для отправки почтового ящика в режиме Compose.
- Добавлена [Office.context.mailbox.item.getComposeTypeAsync:](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.10&preserve-view=true#getComposeTypeAsync_options__callback_)добавляется новая функция, которая получает тип композитного сообщения в режиме Compose.
- Добавлена [Office.context.mailbox.item.isClientSignatureEnabledAsync:](office.context.mailbox.item.md#methods)добавляется новая функция, которая проверяет, включена ли подпись клиента на элементе в режиме Compose.
- Добавлены [Office. MailboxEnums.ActionType:](/javascript/api/outlook/office.mailboxenums.actiontype)Добавляет новый список. Он представляет тип настраиваемого действия в уведомлении.
- Добавлен [Office.MailboxEnums.ComposeType:](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-1.10&preserve-view=true)Добавляет новый список, доступный в режиме Compose.
- Добавлены [Office. MailboxEnums.ItemNotificationMessageType.InsightMessage:](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)добавляет новый `ItemNotificationMessageType` тип в список. Оно представляет сообщение уведомления с пользовательским действием.
- Добавлены [Office. NotificationMessageAction:](/javascript/api/outlook/office.notificationmessageaction)добавляет новый объект, чтобы можно было определить настраиваемые действия для `InsightMessage` уведомления.
- Добавлены [Office. NotificationMessageDetails.actions:](/javascript/api/outlook/office.notificationmessagedetails#actions)добавляет новое свойство, которое позволяет добавлять уведомление `InsightMessage` с помощью настраиваемой меры.

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
