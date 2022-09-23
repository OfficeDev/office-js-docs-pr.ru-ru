---
title: Сравнение поддержки надстроек Outlook в Outlook для Mac
description: Узнайте, как поддержка надстроек в Outlook для Mac сравниваются с другими клиентами Outlook.
ms.date: 09/21/2022
ms.localizationpriority: medium
ms.openlocfilehash: c3f991865921583561e4c2db2132fad3ceba3625
ms.sourcegitcommit: 09bb0b5edd6af03c9822e1742095c7df94735120
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/23/2022
ms.locfileid: "67990415"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Сравнение поддержки надстроек Outlook в Outlook для Mac с другими клиентами Outlook

Вы можете создать и запустить надстройку Outlook в Outlook для Mac так же, как и в других клиентах, включая Outlook в Интернете, Windows, iOS и Android, без настройки JavaScript для каждого клиента. Те же вызовы из надстройки к API JavaScript для Office обычно работают одинаково, за исключением областей, описанных в следующей таблице.

Дополнительные сведения см. в статье [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md).

Дополнительные сведения о поддержке нового пользовательского интерфейса см. в разделе "Поддержка надстроек [" в Outlook для нового пользовательского интерфейса Mac](#add-in-support-in-outlook-on-new-mac-ui).

| Область | Outlook в Интернете, Windows и мобильных устройств | Outlook для Mac |
|:-----|:-----|:-----|
| Поддерживаемые версии файла office.js и схемы манифеста Надстройки Office | Все API в файле office.js и схема версии 1.1. | Все API в файле office.js и схема версии 1.1.<br><br>**ПРИМЕЧАНИЕ**. В Outlook для Mac сохранение собрания поддерживается только в сборке 16.35.308 или более поздней версии. В противном случае `saveAsync` метод завершается сбоем при вызове из собрания в режиме создания. Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745). |
| Экземпляры серии повторяющихся встреч | <ul><li>Можно получить идентификатор элемента и другие свойства основной встречи или экземпляра встречи из серии повторяющихся встреч.</li><li>Можно использовать [mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods), чтобы вывести на экран экземпляр или основную встречу их серии.</li></ul> | <ul><li>Можно получить идентификатор элемента и другие свойства основной встречи, но не экземпляра серии повторяющихся встреч.</li><li>Can display the master appointment of a recurring series. Without the item ID, cannot display an instance of a recurring series.</li></ul> |
| Тип получателя участника встречи | С помощью [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) можно определить тип получателя участника. | `EmailAddressDetails.recipientType` возвращает `undefined` для участников встречи. |
| Строка версии клиентского приложения | Формат строки версии, возвращаемой [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) , зависит от фактического типа клиента. Например:<ul><li>Outlook в Windows: `15.0.4454.1002`</li><li>Outlook в Интернете:`15.0.918.2`</li></ul> |Пример строки версии, возвращаемой `Diagnostics.hostVersion` в Outlook для Mac: `15.0 (140325)` |
| Настраиваемые свойства элемента | Если сеть выходит из строя, надстройка все еще может получить доступ к кэшированным настраиваемым свойствам. | Так как Outlook для Mac не кэширует пользовательские свойства, в случае отключения сети надстройки не смогут получить к ним доступ. |
| Сведения о вложениях | Тип контента и имена вложений в [объекте AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) зависят от типа клиента:<ul><li>Пример `AttachmentDetails.contentType` в формате JSON: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` does not contain any filename extension. As an example, if the attachment is a message that has the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Пример `AttachmentDetails.contentType` в формате JSON: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` always includes a filename extension. Attachments that are mail items have a .eml extension, and appointments have a .ics extension. As an example, if an attachment is an email with the subject "RE: Summer activity", the JSON object that represents the attachment name would be `"name": "RE: Summer activity.eml"`.<p>**Примечание.** Если файл вложен программным образом (например, с помощью надстройки) без расширения, то имя файла в свойстве `AttachmentDetails.name` не будет включать расширение.</p></li></ul> |
| Строка, представляющая часовой пояс в свойствах `dateTimeCreated` и `dateTimeModified` |Пример: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Пример: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Точность времени в свойствах `dateTimeCreated` и `dateTimeModified` | Если надстройка использует приведенный ниже код, то обеспечивается точность до миллисекунд.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| Точность только до секунд. |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>Поддержка надстроек в Outlook в новом пользовательском интерфейсе Mac

Теперь надстройки Outlook поддерживаются в новом пользовательском интерфейсе Mac (доступно в Outlook версии 16.38.506). Сведения о наборах обязательных элементов, которые в настоящее время поддерживаются в новом пользовательском интерфейсе Mac, см. в разделе "Поддержка набора обязательных элементов [API Outlook"](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#outlook-client-support).

Дополнительные сведения о новом пользовательском интерфейсе Mac см. в [Outlook для Mac.](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)

Вы можете определить версию пользовательского интерфейса следующим образом:

**Классический пользовательский интерфейс**

![Классический пользовательский интерфейс на Компьютере Mac.](../images/outlook-on-mac-classic.png)

**Новый пользовательский интерфейс**

![Новый пользовательский интерфейс на Компьютере Mac.](../images/outlook-on-mac-new.png)
