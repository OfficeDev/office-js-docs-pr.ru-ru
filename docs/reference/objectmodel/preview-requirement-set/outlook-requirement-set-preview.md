---
title: Outlook API предварительного просмотра надстройки
description: Функции и API, которые в настоящее время находятся в предварительном Outlook надстройки.
ms.date: 03/15/2022
ms.localizationpriority: medium
ms.openlocfilehash: 714be93351ff67ad49cd07154f145f19949efa68
ms.sourcegitcommit: 856f057a8c9b937bfb37e7d81a6b71dbed4b8ff4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/16/2022
ms.locfileid: "63511274"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook API предварительного просмотра надстройки

Подмножество API Outlook надстройки в API Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.

> [!IMPORTANT]
> Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md). Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке. Не следует указывать этот набор обязательных элементов в манифесте надстройки.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Можно просмотреть функции в Outlook в Интернете, настроив целевой выпуск на [Microsoft 365 клиента](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center). "Настройка доступа к предварительному просмотру" отмечена на этой странице для применимых функций.
>
> Для других функций можно запросить доступ к битам предварительного просмотра для Outlook в Интернете с помощью Microsoft 365 учетной записи, заполнив эту [форму](https://aka.ms/OWAPreview). В этих функциях отмечен "Запрос доступа к предварительному просмотру".

Набор требований предварительного просмотра включает все функции [набора требований 1.11](../requirement-set-1.11/outlook-requirement-set-1.11.md).

## <a name="features-in-preview"></a>Возможности предварительной версии

Ниже перечислены возможности предварительной версии.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Активация надстройки для элементов, защищенных управлением правами на информацию (IRM)

Надстройки теперь могут активироваться в пунктах, защищенных IRM. Чтобы включить эту возможность, `OBJMODEL` администратору клиента необходимо включить право использования, установив в Office параметр **Разрешить** программный доступ. [Дополнительные сведения см. в дополнительных сведениях о](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) правах и описаниях использования.

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

Добавлен новый переумыв `AppointmentSensitivityType` , который представляет параметры чувствительности, доступные при встрече.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

<br>

---

---

### <a name="delay-delivery-time"></a>Время задержки доставки

#### <a name="officecontextmailboxitemdelaydeliverytime"></a>[Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

Добавлено новое свойство, которое возвращает объект, который позволяет управлять датой и временем доставки сообщения в режиме Compose.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officedelaydeliverytime"></a>[Office. DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

Добавлен новый объект, который позволяет управлять датой и временем доставки сообщения в режиме Compose.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

<br>

---

---

### <a name="event-based-activation"></a>Активация на основе событий

Эта функция была выпущена в [наборе требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). Однако дополнительные события теперь доступны в предварительном просмотре. Дополнительные дополнительные ссылки на [поддерживаемые события](../../../outlook/autolaunch.md#supported-events).

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officeaddincommandseventcompletedoptionserrormessage"></a>[Office. AddinCommands.EventCompletedOptions.errorMessage](/javascript/api/office/office.addincommands.eventcompletedoptions?view=outlook-js-preview&preserve-view=true#office-office-addincommands-eventcompletedoptions-errormessage-member)

Добавлено новое свойство для отображения сообщения об ошибке пользователю, если обработанное событие не может продолжать выполняться. Например, обратитесь к [погонам Smart Alerts](../../../outlook/smart-alerts-onmessagesend-walkthrough.md).

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

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

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#office-office-context-officetheme-member)

Добавлена возможность получения темы Office.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)

<br>

---

---

### <a name="shared-mailboxes"></a>Общие почтовые ящики

Поддержка функций для общих папок (то есть доступа делегатов) была выпущена в наборе [требований 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Однако поддержка общих почтовых ящиков теперь доступна в предварительном просмотре. Дополнительные сведения приводятся в статье [Включение сценариев общих папок и общих почтовых ящиков](../../../outlook/delegate-access.md).

**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная), Outlook на Mac

## <a name="see-also"></a>См. также

- [Надстройки Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Примеры кода надстройки Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Начало работы](../../../quickstarts/outlook-quickstart.md)
- [Наборы обязательных элементов и поддерживаемые клиенты](../../requirement-sets/outlook-api-requirement-sets.md)
