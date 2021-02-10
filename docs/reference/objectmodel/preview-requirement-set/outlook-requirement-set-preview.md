---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в предварительной версии для надстройки Outlook.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 92ba3510af0c8b9ebdf9ca4368c889b821a9cb3b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173957"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Подмножество API надстройки Outlook aPI JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройки Outlook.

> [!IMPORTANT]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Вы можете просмотреть функции в Outlook в Интернете, настроив целевой выпуск в [клиенте Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) На этой странице отмечена "Настройка доступа к предварительному просмотру" для применимых функций.
>
> Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook в Интернете с помощью учетной записи Microsoft 365, заполнив и передав [эту форму.](https://aka.ms/OWAPreview) Для этих функций отмечено "Запрашивать предварительный доступ".

Набор предварительных требований включает все функции набора требований [1.9.](../requirement-set-1.9/outlook-requirement-set-1.9.md)

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Активация надстройки для элементов, защищенных с помощью управления правами на управление правами на данные (IRM)

Теперь надстройки могут активироваться для элементов, защищенных с защитой IRM. Чтобы включить эту возможность, администратор клиента должен включить право на использование, задав параметр политики "Разрешить программный доступ" `OBJMODEL` в Office.  Дополнительные [сведения см. в описании](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) и правах на использование.

**Доступно в**: Outlook для Windows, начиная со сборки 13229.10000 (подключенной к подписке На Microsoft 365)

<br>

---

---

### <a name="additional-calendar-properties"></a>Дополнительные свойства календаря

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Добавлен новый объект, который представляет свойство события на весь день встречи в режиме compose.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Добавлен новый объект, который представляет чувствительность встречи в режиме составить.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое представляет, является ли встреча событием на весь день.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое представляет чувствительность встречи.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office.MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Добавлено новое `AppointmentSensitivityType` enum, которое представляет параметры конфиденциальности, доступные для встречи.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Активация на основе событий

Добавлена поддержка функций активации на основе событий в надстройки Outlook. Подробнее [см. в](../../../outlook/autolaunch.md) подстройке "Настройка надстройки Outlook для активации на основе событий".

#### <a name="launchevent-extension-point"></a>[Точка расширения LaunchEvent](../../manifest/extensionpoint.md#launchevent-preview)

Добавлена `LaunchEvent` поддержка точки расширения для манифеста. Он настраивает функции активации на основе событий.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="launchevents-manifest-element"></a>[Элемент манифеста LaunchEvents](../../manifest/launchevents.md)

Добавлен `LaunchEvents` элемент манифеста. Он поддерживает настройку функций активации на основе событий.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="runtimes-manifest-element"></a>[Элемент манифеста runtimes](../../manifest/runtimes.md)

Добавлена поддержка Outlook для `Runtimes` элемента манифеста. Он ссылается на файлы HTML и JavaScript, необходимые для активации на основе событий.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Взаимодействие с интерактивными сообщениями

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

<br>

---

---

### <a name="mail-signature"></a>Подпись почты

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office.context.mailbox.item.body.setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

Добавлена новая функция для объекта, которая добавляет или заменяет подпись в теле `Body` элемента в режиме составить.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office.context.mailbox.item.disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая отключает подпись клиента для отправляемого почтового ящика в режиме составить.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office.context.mailbox.item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

Добавлена новая функция, которая получает тип составить сообщение в режиме составить.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office.context.mailbox.item.isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая проверяет, включена ли подпись клиента для элемента в режиме составить.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

#### <a name="officemailboxenumscomposetype"></a>[Office.MailboxEnums.ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

Добавлено новое `ComposeType` enum, доступное в режиме составить.

**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)

<br>

---

---

### <a name="notification-messages-with-actions"></a>Уведомления с действиями

Эта функция позволяет надстройка включать уведомление с дополнительным действием, кроме действия по умолчанию **"Отклонять".** В современном Outlook в Интернете эта функция доступна только в режиме составить.

#### <a name="officenotificationmessagedetailsactions"></a>[Office.NotificationMessageDetails.actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Добавлено новое свойство, которое позволяет добавить уведомление `InsightMessage` с помощью дополнительного действия.

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

#### <a name="officenotificationmessageaction"></a>[Office.NotificationMessageAction](/javascript/api/outlook/office.notificationmessageaction)

Добавлен новый объект, в котором вы определяете дополнительное действие для `InsightMessage` уведомления.

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

#### <a name="officemailboxenumsactiontype"></a>[Office.MailboxEnums.ActionType](/javascript/api/outlook/office.mailboxenums.actiontype)

Добавлено новое enum `ActionType` .

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office.MailboxEnums.ItemNotificationMessageType.InsightMessage](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Добавлен новый тип `InsightMessage` в `ItemNotificationMessageType` enum.

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

<br>

---

---

### <a name="office-theme"></a>Тема Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Добавлена возможность получения темы Office.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.

**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)

<br>

---

---

### <a name="session-data"></a>Данные сеансов

#### <a name="officesessiondata"></a>[Office.SessionData](/javascript/api/outlook/office.sessiondata)

Добавлен новый объект, который представляет данные сеанса элемента.

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

Добавлено новое свойство для управления данными сеанса элемента в режиме составить.

**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
