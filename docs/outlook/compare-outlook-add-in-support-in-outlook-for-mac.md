---
title: Сравнение поддержки надстроек Outlook в Outlook на компьютерах Mac
description: Узнайте, как сравнить надстройки в Outlook для Mac с другими ведущими приложениями Outlook.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: fd03141fbcaecb88db358101a00681c8a85af382
ms.sourcegitcommit: 71a44405e42b4798a8354f7f96d84548ae7a00f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44280354"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-hosts"></a>Сравнение поддержки надстроек Outlook в Outlook в Mac с другими ведущими приложениями Outlook

Вы можете создавать и запускать надстройку Outlook так же, как и в других узлах, в том числе в Outlook в Интернете, Windows, iOS и Android, без настройки JavaScript для каждого узла. Те же вызовы из надстройки в API JavaScript для Office обычно работают так же, за исключением областей, описанных в следующей таблице.

Дополнительные сведения см. в статье [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md).

| Область | Outlook в Интернете, Windows и мобильных устройствах | Outlook для Mac |
|:-----|:-----|:-----|
| Поддерживаемые версии файла office.js и схемы манифеста Надстройки Office | Все API в файле office.js и схема версии 1.1. | Все API в файле office.js и схема версии 1.1.<br><br>**Примечание**: в Outlook на Mac-адресе только построение 16.35.308 или более поздней версии поддерживает сохранение собрания. В противном случае `saveAsync` метод завершается с ошибкой при вызове из собрания в режиме создания. Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745). |
| Экземпляры серии повторяющихся встреч | <ul><li>Можно получить идентификатор элемента и другие свойства основной встречи или экземпляра встречи из серии повторяющихся встреч.</li><li>Можно использовать [mailbox.displayAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), чтобы вывести на экран экземпляр или основную встречу их серии.</li></ul> | <ul><li>Можно получить идентификатор элемента и другие свойства основной встречи, но не экземпляра серии повторяющихся встреч.</li><li>Можно отобразить основную встречу из серии повторяющихся встреч. Без идентификатора элемента экземпляр серии повторяющихся встреч отобразить невозможно.</li></ul> |
| Тип получателя участника встречи | С помощью [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#recipienttype) можно определить тип получателя участника. | `EmailAddressDetails.recipientType` возвращает `undefined` для участников встречи. |
| Строка версии ведущего приложения | Формат строки версии, возвращаемой [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#hostversion), зависит от фактического типа ведущего приложения. Пример:<ul><li>Outlook в Windows:`15.0.4454.1002`</li><li>Outlook в Интернете:`15.0.918.2`</li></ul> |Пример строки версии, возвращаемой `Diagnostics.hostVersion` в Outlook для Mac:`15.0 (140325)` |
| Настраиваемые свойства элемента | Если сеть выходит из строя, надстройка все еще может получить доступ к кэшированным настраиваемым свойствам. | Так как Outlook на Mac не кэширует настраиваемые свойства, если сеть отключена, надстройки не смогут получить к ним доступ. |
| Сведения о вложениях | Тип контента и имена вложений в объекте [AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) зависят от типа ведущего приложения:<ul><li>Пример `AttachmentDetails.contentType` в формате JSON: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` не содержит расширение файла. Например, если вложение является сообщением с темой "RE: Планы на лето", то объект JSON, представляющий имя этого вложения, будет иметь вид `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Пример `AttachmentDetails.contentType` в формате JSON: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` всегда включает расширение имени файла. Вложения, являющиеся почтовыми элементами, имеют расширение EML, а встречи — расширение ICS. Например, если вложение — сообщение с темой "RE: Планы на лето", имя вложения будет представлено следующим объектом JSON: `"name": "RE: Summer activity.eml"`.<p>**Примечание.** Если файл вложен программным образом (например, с помощью надстройки) без расширения, то имя файла в свойстве `AttachmentDetails.name` не будет включать расширение.</p></li></ul> |
| Строка, представляющая часовой пояс в свойствах `dateTimeCreated` и `dateTimeModified` |Пример: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Пример: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Точность времени в свойствах `dateTimeCreated` и `dateTimeModified` | Если в надстройке используется приведенный ниже код, то обеспечивается точность до миллисекунд:<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| Точность только до секунд. |

