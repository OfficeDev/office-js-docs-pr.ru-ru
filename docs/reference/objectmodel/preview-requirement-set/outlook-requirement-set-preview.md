---
title: Outlook набор требований к предварительному просмотру API надстройки
description: Функции и API, которые в настоящее время находятся в предварительном Outlook надстройки.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: f9d8afc2b4347a8fb13f8ab98a163fb63968123f
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007764"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook набор требований к предварительному просмотру API надстройки

Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!IMPORTANT]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Можно просмотреть функции в Outlook в Интернете, настроив целевой выпуск на [Microsoft 365 клиента.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) "Настройка доступа к предварительному просмотру" отмечена на этой странице для применимых функций.
>
> Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook в Интернете с помощью Microsoft 365 учетной записи, заполнив и подав [эту форму.](https://aka.ms/OWAPreview) В этих функциях отмечен "Запрос доступа к предварительному просмотру".

Набор требований предварительного просмотра включает все функции [набора требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Активация надстройки для элементов, защищенных управлением правами на информацию (IRM)

Надстройки теперь могут активироваться в пунктах, защищенных IRM. Чтобы включить эту возможность, администратору клиента необходимо включить право использования, установив в Office параметр Разрешить программный `OBJMODEL` доступ.  Дополнительные [сведения см. в](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) дополнительных сведениях о правах и описаниях использования.

**Доступно в**: Outlook на Windows, начиная со сборки 13229.10000 (подключен к подписке Microsoft 365)

<br>

---

---

### <a name="additional-calendar-properties"></a>Дополнительные свойства календаря

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Добавлен новый объект, который представляет свойство события на весь день встречи в режиме Compose.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Добавлен новый объект, который представляет чувствительность встречи в режиме Compose.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое представляет, если встреча является событием на весь день.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое представляет чувствительность встречи.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Добавлен новый `AppointmentSensitivityType` переумыв, который представляет параметры чувствительности, доступные при встрече.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Активация на основе событий

Эта функция была выпущена в [наборе требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). Однако дополнительные события теперь доступны в предварительном просмотре. Дополнительные дополнительные ссылки на [поддерживаемые события.](../../../outlook/autolaunch.md#supported-events)

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Взаимодействие с интерактивными сообщениями

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)

<br>

---

---

### <a name="office-theme"></a>Тема Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Добавлена возможность получения темы Office.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

<br>

---

---

### <a name="session-data"></a>Данные сеансов

#### <a name="officesessiondata"></a>[Office. SessionData](/javascript/api/outlook/office.sessiondata)

Добавлен новый объект, который представляет данные сеанса элемента.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

Добавлено новое свойство для управления данными сеанса элемента в режиме Compose.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)

<br>

---

---

### <a name="shared-mailboxes"></a>Общие почтовые ящики

Поддержка функций для общих папок (т. е. доступа делегатов) была выпущена в наборе [требований 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Однако поддержка общих почтовых ящиков теперь доступна в предварительном просмотре. Чтобы узнать больше, обратитесь к [разделу Включить общие папки и сценарии общих почтовых ящиков.](../../../outlook/delegate-access.md)

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
