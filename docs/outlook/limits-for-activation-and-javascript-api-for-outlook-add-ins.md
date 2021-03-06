---
title: Ограничения на активацию и использование API в надстройках Outlook
description: Обратите внимание на определенные правила активации и использования API и учитывайте эти ограничения при реализации своих надстроек.
ms.date: 06/11/2021
localization_priority: Normal
ms.openlocfilehash: 60fab066dadf5c71ab37e907dd749d38f9bb4dde
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348953"
---
# <a name="limits-for-activation-and-javascript-api-for-outlook-add-ins"></a>Ограничения на активацию и API JavaScript для надстроек Outlook

Чтобы предоставить пользователям удобные возможности работы в надстройках Outlook, следует помнить об определенных рекомендациях по активации и использованию интерфейса API и разрабатывать надстройки в соответствии с ними. Эти рекомендации существуют таким образом, что для отдельной надстройки не может потребоваться Exchange Server или Outlook времени на обработку правил активации или вызовов в API javaScript Office, что влияет на общий пользовательский Outlook и другие надстройки. Эти ограничения применяются к разработке правил активации в манифесте надстройки и использованию настраиваемого свойства, параметров роуминга, получателей, Exchange веб-служб (EWS) запросов и ответов, а также асинхронных вызовов.

> [!NOTE]
> Если ваша надстройка работает в полнофункциональном клиенте Outlook, то необходимо также убедиться, что она при этом учитываются ограничения на используемые ресурсы.

## <a name="limits-on-where-add-ins-activate"></a>Где активируются надстройки

Дополнительные сведения о том, где надстройки активируются и не активируются, обратитесь к пунктам почтовых ящиков, доступным в разделе надстройки на странице Outlook обзор надстройки. [](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)

## <a name="limits-for-activation-rules"></a>Ограничения для правил активации

При разработке правил активации надстроек Outlook придерживайтесь следующих рекомендаций.

- Размер манифеста не должен превышать 256 КБ. Если это ограничение будет превышено, установить надстройку Outlook для почтового ящика Exchange будет невозможно.

- Задавайте не более 15 правил активации для надстройки. Если это ограничение будет превышено, установить надстройку не удастся.

- Если правило [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) применяется к тексту сообщения выбранного элемента, следует ожидать, что полнофункциональный клиент Outlook будет применять его только к первому мегабайту текста сообщения, а не ко всему тексту, если его объем превышает это ограничение. Надстройка не будет активирована, если соответствия присутствуют только после первого мегабайта текста сообщения. Если вы считаете такой сценарий вероятным, измените условия активации.

- При использовании регулярных выражений в `ItemHasKnownEntity` [правилах ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) следует помнить о следующих ограничениях и рекомендациях, которые обычно применяются к любому приложению Outlook, а также к выражениям, описанным в таблицах 1, 2 и 3, которые отличаются в зависимости от приложения.
  - Specify up to only five regular expressions in activation rules in a add-in. You cannot install a add-in if you exceed that limit.
  - Укажите регулярные выражения, чтобы результаты, которые вы ожидаете, возвращались методом вызова `getRegExMatches` в течение первых 50 совпадений.
  - Может указывать утверждения с просмотром вперед, но не поддерживает утверждения с просмотром назад `(?<=text)` и отрицательные утверждения с просмотром назад `(?<!text)` в регулярных выражениях.

В таблице 1 перечислены ограничения и описаны различия в поддержке регулярных выражений между Outlook клиентом и Outlook в Интернете или мобильными устройствами. Поддержка не зависит от типа устройства и основного текста элемента.

**Таблица 1. Основные различия в поддержке регулярных выражений**

|Полнофункциональный клиент Outlook|Outlook в Интернете или на мобильных устройствах|
|:-----|:-----|
|Использует обработчик регулярных выражений C++, который предоставляется как часть библиотеки стандартных шаблонов Visual Studio. Этот обработчик выполняет компиляцию в соответствии со стандартами ECMAScript 5. |Использует оценку регулярных выражений, которая является частью JavaScript, предоставляется браузером и поддерживает расширенный набор стандартов ECMAScript 5.|
|Из-за различных регексных движков следует ожидать, что regex, включаемый в настраиваемый класс символов на основе предопределяемого класса символов, может возвращать различные результаты в Outlook клиенте, чем на Outlook в Интернете или мобильных устройствах.<br/><br/>Например, регулярное выражение `[\s\S]{0,100}` соответствует любому количеству (от 0 до 100) символов, включая пробелы. Этот regex возвращает различные результаты в Outlook клиента, чем Outlook в Интернете и мобильных устройств.<br/><br/>Чтобы обойти эту проблему, следует переписать его в виде `(\s\|\S){0,100}`. Такое регулярное выражение соответствует любому количеству символов (от 0 до 100), включая пробелы.<br/><br/>Необходимо тщательно протестировать каждый regex на каждом Outlook клиенте, и если regex возвращает разные результаты, переопиши regex. |Необходимо тщательно протестировать каждый regex на каждом Outlook клиенте, и если regex возвращает разные результаты, переопиши regex.|
|По умолчанию продолжительность оценки всех регулярных выражений для надстройки ограничивается 1 секундой. При превышении этого ограничения выполняется повторная оценка до 3 раз. Помимо ограничения переоплаты, Outlook клиент отключает надстройку от работы для одного и того же почтового ящика в любом Outlook клиентах.<br/><br/>Администраторы могут переопределять эти ограничения оценки с помощью `OutlookActivationAlertThreshold` ключей реестра и `OutlookActivationManagerRetryLimit` реестра.|Не поддерживаются такие же параметры реестра и мониторинга ресурсов, как в полнофункциональном клиенте Outlook. Но надстройки с регулярными выражениями, которые требуют чрезмерного времени на оценку Outlook клиента, отключены для одного и того же почтового ящика для всех Outlook клиентов.|

В таблице 2 перечислены ограничения и описаны различия частей основного текста элемента, к которым каждое приложение Outlook применяет регулярные выражения. Некоторые из этих ограничений зависят от типа устройства и основного текста элемента, если регулярное выражение применяется к основному тексту элемента.

**Таблица 2. Ограничения на размер оцениваемого содержания элемента**

||Полнофункциональный клиент Outlook|Outlook на мобильных устройствах|Outlook в Интернете|
|:-----|:-----|:-----|:-----|
|**Форм-фактор**|Любое поддерживаемое устройство|Смартфоны Android, iPad или iPhone|Все поддерживаемые устройства, кроме смартфонов Android, iPad и iPhone|
|**Основной текст элемента в виде обычного текста**|Регулярное выражение применяется к первому мегабайту данных в основном тексте. К остальной части основного текста свыше этого ограничения регулярное выражение не применяется.|Надстройка активируется, только если основной текст сообщения содержит < 16 000 символов.|Надстройка активируется, только если основной текст сообщения содержит < 500 000 символов.|
|**Основной текст элемента в формате HTML**|Регулярное выражение применяется к первым 512 КБ данных в основном тексте. К остальной части основного текста свыше этого ограничения регулярное выражение не применяется. (Фактическое количество знаков зависит от кодировки, в которой может использоваться от 1 до 4 байтов на каждый знак.)|Регулярное выражение применяется к первым 64 000 знаков (включая знаки HTML-тегов). К остальной части основного текста свыше этого ограничения регулярное выражение не применяется.|Надстройка активируется, только если основной текст сообщения содержит < 500 000 символов.|

В таблице 3 перечислены ограничения и описаны различия в совпадениях, которые Outlook возвращает клиент после оценки регулярного выражения. Поддержка не зависит от конкретного типа устройства, но может зависеть от типа основного текста элемента, если регулярное выражение применяется к основному тексту элемента.

**Таблица 3. Ограничения на возвращаемые соответствия**

||Полнофункциональный клиент Outlook|Outlook в Интернете или на мобильных устройствах|
|:-----|:-----|:-----|
|**Порядок возвращаемых соответствий**|Предположим, что возвращает совпадения для одного и того же регулярного выражения, применяемого на одном и том же элементе, отличаются в Outlook клиенте, чем Outlook в Интернете `getRegExMatches` или мобильных устройствах.|Предположим, возвращает совпадения в другом `getRegExMatches` порядке в Outlook клиенте, чем Outlook в Интернете или мобильных устройствах.|
|**Основной текст элемента в виде обычного текста**|`getRegExMatches` возвращает любые совпадения, которые имеют значение до 1536 символов (1,5 КБ), максимум для 50 совпадений.<br/><br/>**Примечание.** `getRegExMatches` Не возвращает совпадения в определенном порядке в возвращаемом массиве. В общем, предположить, что порядок совпадений в Outlook клиенте для одного и того же регулярного выражения, применяемого на одном и том же элементе, отличается от порядка в Outlook в Интернете и мобильных устройствах.|`getRegExMatches` возвращает любые совпадения, которые имеют до 3072 (3 КБ) символов, не более 50 совпадений.|
|**Основной текст элемента в формате HTML**|`getRegExMatches` возвращает любые совпадения, которые имеют до 3072 (3 КБ) символов, не более 50 совпадений.<br/> <br/> **Примечание.** `getRegExMatches` Не возвращает совпадения в определенном порядке в возвращаемом массиве. В общем, предположить, что порядок совпадений в Outlook клиенте для одного и того же регулярного выражения, применяемого на одном и том же элементе, отличается от порядка в Outlook в Интернете и мобильных устройствах.|`getRegExMatches` возвращает любые совпадения, которые имеют до 3072 (3 КБ) символов, не более 50 совпадений.|

## <a name="limits-for-javascript-api"></a>Ограничения для API JavaScript

Помимо предыдущих правил активации, каждый клиент Outlook применяет определенные ограничения в объектной модели JavaScript, как описано в таблице 4.

**Таблица 4. Ограничения для получения или набора определенных данных с Office API JavaScript**

|Функция|Ограничение|Связанный API|Описание|
|:-----|:-----|:-----|:-----|
|Настраиваемые свойства|2500 символов|Объект [CustomProperties](/javascript/api/outlook/office.CustomProperties)<br/> <br/>Метод [item.loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)|Ограничение для всех настраиваемых свойств элемента встречи или сообщения. Все Outlook возвращают ошибку, если общий размер всех пользовательских свойств надстройки превышает этот предел.|
|Параметры перемещения|32 КБ символов|Объект [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)<br/><br/> Свойство [context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md#properties)|Ограничение для всех параметров перемещения надстройки. Все Outlook возвращают ошибку, если параметры превышают этот предел.|
|Извлечение известных сущностей|2000 символов|Метод [item.getEntities](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)<br/> <br/>Метод [item.getEntitiesByType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)<br/> <br/>Метод [item.getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)|Ограничение для Exchange Server для извлечения известных сущностей в основном тексте элемента. Exchange Server игнорирует сущности сверх этого предела. Обратите внимание, что это ограничение не зависит от того, использует ли надстройка `ItemHasKnownEntity` правило.|
|Веб-службы Exchange|1 МБ символов|Метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Ограничение для запроса или ответа на `Mailbox.makeEwsRequestAsync` вызов.|
|Получатели|100 получателей|Свойство [item.requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)<br/> <br/>Свойство [item.optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)<br/> <br/>Свойство [item.to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)<br/> <br/>Свойство [item.cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)<br/> <br/>Метод [Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)<br/> <br/>Метод [Recipient.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)<br/> <br/>Метод [Recipient.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)|Ограничение для получателей, указанных в каждом свойстве.|
|Отображаемое имя|255 символов|Свойство [EmailAddressDetails.displayName](/javascript/api/outlook/office.emailaddressdetails#displayname)<br/><br/> Объект [Recipients](/javascript/api/outlook/office.Recipients)<br/><br/> `item.requiredAttendees` свойство<br/><br/> `item.optionalAttendees` свойство <br/><br/>`item.to` свойство <br/><br/>`item.cc` свойство|Ограничение длины отображаемого имени в сообщении или встрече.|
|Настройка темы|255 символов|Метод [mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)<br/><br/> Метод [Subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)|Ограничение для темы в форме новой встречи или для настройки темы встречи или сообщения.|
|Установка расположения|255 символов|Метод [Location.setAsync](/javascript/api/outlook/office.Location#setasync-location--options--callback-)|Ограничение для установки расположении встречи или приглашения на собрание.|
|Текст в форме новой встречи|32 КБ символов|`Mailbox.displayNewAppointmentForm` метод|Ограничение текста в форме новой встречи.|
|Отображение текста существующего элемента|32 КБ символов|Метод [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)<br/><br/> Метод [mailbox.displayMessageForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Для Outlook в Интернете и мобильных устройств: ограничение для тела в существующей форме встречи или сообщения.|
|Настройка текста|1 МБ символов|Метод [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-)<br/> <br/>[Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-)<br/><br/>Метод [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)|Ограничение для установки текста элемента встречи или сообщения.|
|Число вложений|499 файлов на Outlook в Интернете и мобильных устройствах|Метод [item.addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)|Ограничение количества файлов, которые можно вложить в отправляемый элемент. Outlook в Интернете и мобильные устройства обычно ограничивают присоединение до 499 файлов с помощью пользовательского интерфейса и `addFileAttachmentAsync` . В полнофункциональном клиенте Outlook нет определенного ограничения на количество вложенных файлов. Однако все Outlook соблюдают ограничение размера вложений, с Exchange Server с помощью вложений. См. следующую строку — "Размер вложений".|
|Размер вложений|В зависимости от Exchange Server|`item.addFileAttachmentAsync` метод|Существует ограничение на размер всех вложений элемента, которое администратор может настроить на сервере Exchange Server для почтового ящика пользователя. В полнофункциональном клиенте Outlook это ограничивает количество вложений в элементе. Для Outlook в Интернете и мобильных устройств меньшее из этих двух ограничений — количество вложений и размер всех вложений — ограничивает фактические вложения для элемента.|
|Имя файла вложения|255 символов|`item.addFileAttachmentAsync` метод|Ограничение длины имени файла вложения, добавляемого в элемент.|
|URI вложения|2048 символов|`item.addFileAttachmentAsync` метод|Ограничение URI имени файла, добавляемого в элемент как вложение.|
|Идентификатор вложения|100 символов|Метод [item.addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)<br/><br/> Метод [item.removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)|Ограничение длины идентификатора вложения, добавляемого в элемент или удаляемого из него.|
|Асинхронные вызовы|3 вызова|`item.addFileAttachmentAsync` метод<br/><br/>`item.addItemAttachmentAsync` метод<br/><br/><br/>`item.removeAttachmentAsync` метод<br/><br/> Метод [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-)<br/><br/>`Body.prependAsync` метод<br/><br/>`Body.setSelectedDataAsync` метод<br/><br/> Метод [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-)<br/><br/><br/> Метод [item.LoadCustomPropertiesAysnc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)<br/><br/><br/> Метод [Location.getAsync](/javascript/api/outlook/office.Location#getasync-options--callback-)<br/><br/>`Location.setAsync` метод<br/><br/> Метод [mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)<br/><br/> Метод [mailbox.getUserIdentityTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)<br/><br/> Метод [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)<br/><br/>`Recipients.addAsync` метод<br/><br/> Метод [Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-)<br/><br/>`Recipients.setAsync` метод<br/><br/> Метод [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-)<br/><br/> Метод [Subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-)<br/><br/>`Subject.setAsync` метод<br/><br/> Метод [Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-)<br/><br/> Метод [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)|Для Outlook в Интернете или мобильных устройств: ограничение количества одновременных асинхронных вызовов в любое время, так как браузеры позволяют только ограниченное количество асинхронных вызовов на серверы. |

## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)
- [Конфиденциальность, разрешения и безопасность для надстроек Outlook](../concepts/privacy-and-security.md)
