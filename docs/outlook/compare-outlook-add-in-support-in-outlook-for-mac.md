---
title: Сравнение Outlook надстроек в Outlook Mac
description: Узнайте, как поддержка надстроек в Outlook Mac сравниваются с другими Outlook клиентами.
ms.date: 06/09/2022
ms.localizationpriority: medium
ms.openlocfilehash: 36a10f0454bebf3f069464277c7eb2a8a18f42b7
ms.sourcegitcommit: 2eeb0423a793b3a6db8a665d9ae6bcb10e867be3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/10/2022
ms.locfileid: "66019607"
---
# <a name="compare-outlook-add-in-support-in-outlook-on-mac-with-other-outlook-clients"></a>Сравнение Outlook надстроек в Outlook Mac с другими Outlook клиентами

Вы можете создать и запустить надстройку Outlook в Outlook на Mac так же, как и на других клиентах, включая Outlook в Интернете, Windows, iOS и Android, не настраивая JavaScript для каждого клиента. Те же вызовы из надстройки к API JavaScript Office обычно работают одинаково, за исключением областей, описанных в следующей таблице.

Дополнительные сведения см. в статье [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md).

Дополнительные сведения о поддержке нового пользовательского интерфейса см. в разделе "Поддержка надстроек[" Outlook в новом пользовательском интерфейсе Mac](#add-in-support-in-outlook-on-new-mac-ui).

| Область | Outlook в Интернете, Windows и мобильных устройств | Outlook для Mac |
|:-----|:-----|:-----|
| Поддерживаемые версии файла office.js и схемы манифеста Надстройки Office | Все API в файле office.js и схема версии 1.1. | Все API в файле office.js и схема версии 1.1.<br><br>**ПРИМЕЧАНИЕ**. В Outlook Mac только сборка 16.35.308 или более поздней поддерживает сохранение собрания. В противном случае `saveAsync` метод завершается сбоем при вызове из собрания в режиме создания. Временное решение представлено в статье [Не удается сохранить встречу как черновик в Outlook для Mac с помощью API JS для Office](https://support.microsoft.com/help/4505745). |
| Экземпляры серии повторяющихся встреч | <ul><li>Можно получить идентификатор элемента и другие свойства основной встречи или экземпляра встречи из серии повторяющихся встреч.</li><li>Можно использовать [mailbox.displayAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods), чтобы вывести на экран экземпляр или основную встречу их серии.</li></ul> | <ul><li>Можно получить идентификатор элемента и другие свойства основной встречи, но не экземпляра серии повторяющихся встреч.</li><li>Можно отобразить основную встречу из серии повторяющихся встреч. Без идентификатора элемента экземпляр серии повторяющихся встреч отобразить невозможно.</li></ul> |
| Тип получателя участника встречи | С помощью [EmailAddressDetails.recipientType](/javascript/api/outlook/office.emailaddressdetails#outlook-office-emailaddressdetails-recipienttype-member) можно определить тип получателя участника. | `EmailAddressDetails.recipientType` возвращает `undefined` для участников встречи. |
| Строка версии клиентского приложения | Формат строки версии, возвращаемой [diagnostics.hostVersion](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostversion-member) , зависит от фактического типа клиента. Например,<ul><li>Outlook на Windows:`15.0.4454.1002`</li><li>Outlook в Интернете:`15.0.918.2`</li></ul> |Пример строки версии, возвращаемой `Diagnostics.hostVersion` Outlook Mac:`15.0 (140325)` |
| Настраиваемые свойства элемента | Если сеть выходит из строя, надстройка все еще может получить доступ к кэшированным настраиваемым свойствам. | Так Outlook Mac не кэширует настраиваемые свойства, если сеть выходит из строя, надстройки не смогут получить к ним доступ. |
| Сведения о вложениях | Тип контента и имена вложений в [объекте AttachmentDetails](/javascript/api/outlook/office.attachmentdetails) зависят от типа клиента:<ul><li>Пример `AttachmentDetails.contentType` в формате JSON: `"contentType": "image/x-png"`. </li><li>`AttachmentDetails.name` не содержит расширение файла. Например, если вложение является сообщением с темой "RE: Планы на лето", то объект JSON, представляющий имя этого вложения, будет иметь вид `"name": "RE: Summer activity"`.</li></ul> | <ul><li>Пример `AttachmentDetails.contentType` в формате JSON: `"contentType" "image/png"`</li><li>`AttachmentDetails.name` всегда включает расширение имени файла. Вложения, являющиеся почтовыми элементами, имеют расширение EML, а встречи — расширение ICS. Например, если вложение — сообщение с темой "RE: Планы на лето", имя вложения будет представлено следующим объектом JSON: `"name": "RE: Summer activity.eml"`.<p>**Примечание.** Если файл вложен программным образом (например, с помощью надстройки) без расширения, то имя файла в свойстве `AttachmentDetails.name` не будет включать расширение.</p></li></ul> |
| Строка, представляющая часовой пояс в свойствах `dateTimeCreated` и `dateTimeModified` |Пример: `Thu Mar 13 2014 14:09:11 GMT+0800 (China Standard Time)` | Пример: `Thu Mar 13 2014 14:09:11 GMT+0800 (CST)` |
| Точность времени в свойствах `dateTimeCreated` и `dateTimeModified` | Если надстройка использует приведенный ниже код, то обеспечивается точность до миллисекунд.<br/>`JSON.stringify(Office.context.mailbox.item, null, 4);`| Точность только до секунд. |

## <a name="add-in-support-in-outlook-on-new-mac-ui"></a>Поддержка надстроек в Outlook пользовательском интерфейсе Mac

Outlook надстройки теперь поддерживаются в новом пользовательском интерфейсе Mac (доступно из Outlook версии 16.38.506) до набора обязательных элементов 1.10. Однако следующие наборы требований и функции **пока не** поддерживаются.

- Набор обязательных элементов API 1.11

Дополнительные сведения о новом пользовательском интерфейсе Mac см. в [Outlook для Mac.](https://support.microsoft.com/office/6283be54-e74d-434e-babb-b70cefc77439)

Вы можете определить версию пользовательского интерфейса следующим образом:

**Классический пользовательский интерфейс**

![Классический пользовательский интерфейс на Компьютере Mac.](../images/outlook-on-mac-classic.png)

**Новый пользовательский интерфейс**

![Новый пользовательский интерфейс на Компьютере Mac.](../images/outlook-on-mac-new.png)
