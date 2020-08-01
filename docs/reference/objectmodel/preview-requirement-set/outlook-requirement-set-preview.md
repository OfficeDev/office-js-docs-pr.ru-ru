---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в режиме предварительной версии для надстроек Outlook.
ms.date: 07/22/2020
localization_priority: Normal
ms.openlocfilehash: 9a2ca9a71498c954d9864f2b967b5fb240f8c506
ms.sourcegitcommit: 7d5407d3900d2ad1feae79a4bc038afe50568be0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2020
ms.locfileid: "46530480"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!IMPORTANT]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Вы можете предварительно просмотреть функции в Outlook в Интернете, [настроив целевой выпуск на клиенте Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center). "Настройка предварительного доступа" отмечено на этой странице в соответствующих возможностях.
>
> Для других функций вы можете запросить доступ к предварительной версии BITS для Outlook в Интернете, используя свою учетную запись Microsoft 365, заполнив и отправив [эту форму](https://aka.ms/OWAPreview). В этих функциях указано "запросить доступ к предварительному доступу".

Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

### <a name="additional-calendar-properties"></a>Дополнительные свойства календаря

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

Добавлен новый объект, представляющий свойство события "целый день" для встречи в режиме создания.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

Добавлен новый объект, представляющий чувствительность встречи в режиме создания.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office. Context. Mailbox. Item. Исаллдайевент](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое указывает, является ли встреча событием на целый день.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. Context. Mailbox. Item. чувствительность](office.context.mailbox.item.md#properties)

Добавлено новое свойство, представляющее чувствительность встречи.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums. Аппоинтментсенситивититипе](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

Добавлено новое перечисление `AppointmentSensitivityType` , представляющее параметры конфиденциальности, доступные для встречи.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

<br>

---

---

### <a name="append-on-send"></a>Добавление при отправке

Для получения сведений об использовании функции "присоединение к отправке", ознакомьтесь со статьей [Реализация добавления при отправке в надстройке Outlook](../../../outlook/append-on-send.md).

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office. Context. Mailbox. Item. Body. Аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="async-versions-of-display-apis"></a>Асинхронные версии `display` API

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[Office. Context. Mailbox. Дисплайаппоинтментформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

Добавлена новая функция для `Mailbox` объекта, отображающего существующую встречу. Это асинхронная версия `displayAppointmentForm` метода.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[Office. Context. Mailbox. Дисплаймессажеформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

Добавлена новая функция для `Mailbox` объекта, отображающего существующее сообщение. Это асинхронная версия `displayMessageForm` метода.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[Office. Context. Mailbox. Дисплайневаппоинтментформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

Добавлена новая функция для `Mailbox` объекта, отображающего новую форму встречи. Это асинхронная версия `displayNewAppointmentForm` метода.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[Office. Context. Mailbox. Дисплайневмессажеформасинк](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

Добавлена новая функция для `Mailbox` объекта, отображающего форму нового сообщения. Это асинхронная версия `displayNewMessageForm` метода.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[Office. Context. Mailbox. Item. Дисплайрепляллформасинк](office.context.mailbox.item.md#methods)

Добавлена новая функция для `Item` объекта, отображающего форму "ответить всем" в режиме чтения. Это асинхронная версия `displayReplyAllForm` метода.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[Office. Context. Mailbox. Item. Дисплайреплиформасинк](office.context.mailbox.item.md#methods)

Добавлена новая функция для `Item` объекта, отображающего форму "Reply" в режиме чтения. Это асинхронная версия `displayReplyForm` метода.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Активация на основе событий

Добавлена поддержка функций активации на основе событий в надстройках Outlook. Чтобы узнать больше, ознакомьтесь со статьей [Настройка надстройки Outlook для активации на основе событий](../../../outlook/autolaunch.md) .

#### <a name="launchevent-extension-point"></a>[Точка расширения Лаунчевент](../../manifest/extensionpoint.md#launchevent-preview)

Добавлена `LaunchEvent` Поддержка точек расширения для манифеста. Он настраивает функции активации на основе событий.

**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))

#### <a name="launchevents-manifest-element"></a>[Элемент манифеста Лаунчевентс](../../manifest/launchevents.md)

Добавлен `LaunchEvents` элемент для манифеста. Он поддерживает настройку функций активации на основе событий.

**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))

#### <a name="runtimes-manifest-element"></a>[Элемент манифеста среды выполнения](../../manifest/runtimes.md)

Добавлена поддержка Outlook для `Runtimes` элемента manifest. Он ссылается на HTML-и JavaScript-файлы, необходимые для функции активации на основе событий.

**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))

<br>

---

---

### <a name="get-all-custom-properties"></a>Получение всех настраиваемых свойств

#### <a name="custompropertiesgetall"></a>[CustomProperties. Жеталл](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

Добавлена новая функция для `CustomProperties` объекта, который получает все настраиваемые свойства.

**Доступно в**: Outlook в Windows (подключенном к подписке на Microsoft 365), Outlook в Интернете (современная), Outlook на Mac (подключено к подписке Microsoft 365), Outlook на Android, Outlook на iOS

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Взаимодействие с интерактивными сообщениями

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (классическая)

<br>

---

---

### <a name="mail-signature"></a>Подпись почты

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office. Context. Mailbox. Item. Жеткомпосетипеасинк](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк](office.context.mailbox.item.md#methods)

Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

#### <a name="officemailboxenumscomposetype"></a>[Office. MailboxEnums. Компосетипе](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

Добавлено новое перечисление `ComposeType` , доступное в режиме создания.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))

<br>

---

---

### <a name="notification-messages-with-actions"></a>Сообщения уведомления с действиями

Эта функция позволяет надстройке включать сообщение уведомления с дополнительным **действием, кроме действия по** умолчанию.

#### <a name="officenotificationmessagedetailsactions"></a>[Office. NotificationMessageDetails. Actions](/javascript/api/outlook/office.notificationmessagedetails#actions)

Добавлено новое свойство, которое позволяет добавить `InsightMessage` уведомление с дополнительным действием.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)

#### <a name="officenotificationmessageaction"></a>[Office. Нотификатионмессажеактион](/javascript/api/outlook/office.notificationmessageaction)

Добавлен новый объект, в котором определяется дополнительное действие для `InsightMessage` уведомления.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)

#### <a name="officemailboxenumsactiontype"></a>[Office. MailboxEnums.](/javascript/api/outlook/office.mailboxenums.actiontype)

Добавлено новое перечисление `ActionType` .

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[Office. MailboxEnums. Итемнотификатионмессажетипе. Инсигхтмессаже](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

Добавлен новый тип `InsightMessage` в `ItemNotificationMessageType` перечисление.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)

<br>

---

---

### <a name="office-theme"></a>Тема Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Добавлена возможность получения темы Office.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.

**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)

<br>

---

---

### <a name="single-sign-on-sso"></a>Единый вход

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.

**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook на Mac (подключен к подписке Microsoft 365), Outlook в Интернете (современный), Outlook в Интернете (классическая)

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
