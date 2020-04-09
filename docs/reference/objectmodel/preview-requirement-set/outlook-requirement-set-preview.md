---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в предварительной версии для надстроек Outlook и API JavaScript для Office.
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: acc19c81f929596b0bd5622e696c1988cf31ee5c
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185416"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Предварительная версия набора обязательных элементов API для надстройки Outlook

Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.

> [!IMPORTANT]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

### <a name="additional-calendar-properties"></a>Дополнительные свойства календаря

#### <a name="isalldayevent"></a>[исаллдайевент](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

Добавлен новый объект, представляющий свойство события "целый день" для встречи в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

Добавлен новый объект, представляющий чувствительность встречи в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office. Context. Mailbox. Item. Исаллдайевент](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое указывает, является ли встреча событием на целый день.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. Context. Mailbox. Item. чувствительность](office.context.mailbox.item.md#properties)

Добавлено новое свойство, представляющее чувствительность встречи.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums. Аппоинтментсенситивититипе](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

Добавлено новое перечисление `AppointmentSensitivityType` , представляющее параметры конфиденциальности, доступные для встречи.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

<br>

---

---

### <a name="append-on-send"></a>Добавление при отправке

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office. Context. Mailbox. Item. Body. Аппендонсендасинк](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.

**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.

**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Взаимодействие с интерактивными сообщениями

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)

<br>

---

---

### <a name="mail-signature"></a>Подпись почты

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office. Context. Mailbox. Item. Жеткомпосетипеасинк](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк](office.context.mailbox.item.md#methods)

Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officemailboxenumscomposetype"></a>[Office. MailboxEnums. Компосетипе](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

Добавлено новое перечисление `ComposeType` , доступное в режиме создания.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

<br>

---

---

### <a name="office-theme"></a>Тема Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Добавлена возможность получения темы Office.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)

<br>

---

---

### <a name="online-meeting-provider-integration"></a>Интеграция поставщика собраний по сети

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[Точка расширения Мобилеонлинемитингкоммандсурфаце](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

Добавлена `MobileOnlineMeetingCommandSurface` точка расширения для манифеста. Он определяет интеграцию собраний по сети.

**Доступно в**: Outlook на Android (подключено к подписке Office 365)

<br>

---

---

### <a name="sso"></a>Единый вход

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.

**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
