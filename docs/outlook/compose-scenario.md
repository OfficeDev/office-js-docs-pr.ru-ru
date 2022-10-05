---
title: Разработка надстроек Outlook для форм создания
description: Узнайте о сценариях и возможностях надстроек Outlook для форм создания.
ms.date: 10/03/2022
ms.localizationpriority: high
ms.openlocfilehash: ef81b21eaa0bc63a5bf38757cb188e8850ade443
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467253"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a>Разработка надстроек Outlook для форм создания

Вы можете создавать надстройки compose, которые являются надстройки Outlook, активированные в формах создания. В отличие от надстроек чтения (надстройки Outlook, которые активируются в режиме чтения, когда пользователь просматривает сообщение или встречу), надстройки создания доступны в следующих сценариях пользователя.

- Создание сообщения, приглашения на собрание или встречи в отдельной форме.

- Просмотр или редактирование существующих встречи или собрания, организованных пользователем.

   > [!NOTE]
   > If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available. Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.

- Создание ответа на сообщение (встроенного или в отдельной форме).

- Изменение ответа (**Принять**, **Под вопросом** или **Отклонить**) на приглашение на собрание или элемент собрания.

- Предложение нового времени для элемента собрания.

- Пересылка или ответ на приглашение на собрание или элемент собрания.

In each of these compose scenarios, any add-in command buttons defined by the add-in are shown. For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in. The following figure shows add-in commands in a compose form.

![Форма создания элемента Outlook с командами надстройки](../images/compose-form-commands.png)

На рисунке ниже показана область выбора надстроек, включающая две надстройки создания, в которых не реализованы команды. Она активируется при создании встроенного ответа в Outlook.

![Почтовое приложение, содержащее шаблоны, которое активировано в форме создания.](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a>Типы надстроек, доступные в режиме создания

Надстройки создания реализуются в виде [команд надстроек Outlook](add-in-commands-for-outlook.md). Чтобы надстройки активировались при создании писем или ответов на приглашения на собрания, в манифест включается [точка расширения MessageComposeCommandSurface](/javascript/api/manifest/extensionpoint#messagecomposecommandsurface). Чтобы надстройки активировались при создании или редактировании встреч или собраний, организованных пользователем, добавляется [точка расширения AppointmentOrganizerCommandSurface](/javascript/api/manifest/extensionpoint#appointmentorganizercommandsurface).

> [!NOTE]
> На серверах или клиентах, не поддерживающих команды надстроек, используются [правила активации](activation-rules.md), указанные в элементе [Rule](/javascript/api/manifest/rule), содержащемся в элементе [OfficeApp](/javascript/api/manifest/officeapp). Если надстройка не разрабатывается специально для устаревших клиентов и серверов, в ней следует использовать команды надстроек.

## <a name="api-features-available-to-compose-add-ins"></a>Функции API, доступные надстройкам создания

- [Добавление и удаление вложений в форме создания Outlook](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [Просмотр и изменение данных элемента в форме создания элементов Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook](get-set-or-add-recipients.md)
- [Просмотр или изменение темы при создании встречи или сообщения в Outlook](get-or-set-the-subject.md)
- [Вставка данных в текст при создании встречи или сообщения в Outlook](insert-data-in-the-body.md)
- [Просмотр или изменение расположения при создании встречи в Outlook](get-or-set-the-location-of-an-appointment.md)
- [Просмотр или изменение времени при создании встречи в Outlook](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a>См. также

- [Начало работы с надстройками Outlook для Office](../quickstarts/outlook-quickstart.md)
