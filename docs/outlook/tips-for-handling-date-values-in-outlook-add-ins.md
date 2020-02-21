---
title: Обработка значений дат в надстройках Outlook
description: В интерфейсе API JavaScript для Office для хранения и извлечения даты и времени используется преимущественно объект JavaScript Date.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5718839ebda433df6fb14886da34d734f81eb5f2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166640"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Советы по обработке значений дат в надстройках Outlook

В интерфейсе API JavaScript для Office для хранения и извлечения даты и времени используется преимущественно объект JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp). 

Такой объект **Date** обеспечивает методы [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) и [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), которые возвращают запрос значения даты и времени в формате всемирного координированного времени (UTC).

Объект **Date** обеспечивает также другие методы, например [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) и [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), которые возвращают запрос даты или времени по "местному времени".

Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, установленных на клиентских компьютерах под управлением Windows, при вызове метода JavaScript **getDate**, возвращается дата на основе часового пояса, установленного в операционной системе Windows на клиентском компьютере.

В указанном ниже примере создается объект **Date** с именем `myLocalDate` для местного времени, а затем вызывается метод **toUTCString** для преобразования этой даты в строку формата UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Объект JavaScript **Date** можно использовать для получения значения даты или времени на основе времени в формате UTC или местного времени клиентского компьютера, однако объект **Date** ограничен тем, что не поддерживает метод возвращения значения даты или времени для определенного часового пояса. Например, если клиентский компьютер настроен для использования восточного поясного времени (EST), нет ни одного метода **Date**, который позволит получить значение часа в часовом поясе, отличном от EST или UTC, например значение часа по тихоокеанскому времени (PST).


## <a name="date-related-features-for-outlook-add-ins"></a>Функции надстроек Outlook, связанные с датой

При использовании API JavaScript для Office для обработки значений даты и времени в надстройках Outlook, выполняемых в полнофункциональном клиенте Outlook, а также в Outlook в Интернете или на мобильных устройствах, упомянутые выше ограничения JavaScript не имеют значения.


### <a name="time-zones-for-outlook-clients"></a>Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.

|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. В большинстве браузеров для отображения значений даты и времени объекта JavaScript **Date** используется часовой пояс клиентского компьютера.<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. Если пользователь изменяет часовой пояс на клиентском компьютере в Mac, Outlook в MAC-адресе будет предлагать пользователю обновить часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователь задает это значение часового пояса (и предпочитаемый язык), когда он впервые выполняет вход в Outlook в Интернете или на мобильных устройствах. <br/><br/>В Outlook в Интернете и на мобильных устройствах этот часовой пояс используется для отображения значений даты и времени в пользовательском интерфейсе.|

Так как расширенный клиент Outlook использует часовой пояс клиентского компьютера, а пользовательский интерфейс Outlook в Интернете и на мобильных устройствах использует часовой пояс центра администрирования Exchange, местное время для той же надстройки, установленной для того же почтового ящика, может отличаться при работе в расширенной клие Outlook NT и в Outlook в Интернете или на мобильных устройствах. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.


### <a name="date-related-api"></a>API, связанный с датами

Ниже приведены свойства и методы в интерфейсе JavaScript API для Office, которые поддерживают функциональные возможности, связанные с датами.

**Элемент API**|**Представление часового пояса**|**Пример в расширенном клиенте Outlook**|**Пример в Outlook в Интернете или на мобильных устройствах**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. В Outlook в Интернете и мобильных устройствах это свойство возвращает часовой пояс центра администрирования Exchange. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Каждое из этих свойств возвращает объект JavaScript **Date**. Это значение **даты** отображается в формате UTC, как показано в следующем примере — `myUTCDate` имеет то же значение, что и в расширенном клиенте Outlook, Outlook в Интернете и на мобильных устройствах.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Однако вызов `myDate.getDate` возвращает значение даты в часовом поясе клиентского компьютера, которое согласуется с часовым поясом, используемым для отображения значений даты и времени в пользовательском интерфейсе Outlook с расширенными возможностями, но может отличаться от часовых поясов, которые Outlook в Интернете и мобильные устройства используют в своем пользовательском интерфейсе.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Для каждого из параметров _Start_ и _End_ требуется объект JavaScript **Date**. Аргументы должны быть правильно заданы в формате UTC независимо от часового пояса, используемого в пользовательском интерфейсе в расширенном клиенте Outlook или в Интернете или на мобильных устройствах.|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>Вспомогательные методы для сценариев, связанных с датами


Как описано в предыдущих разделах, так как "Местное время" для пользователя в Outlook в Интернете или мобильных устройствах может различаться в расширенном клиенте Outlook, но объект JavaScript **Date** поддерживает преобразование только в часовой пояс клиентского компьютера или в формате UTC, API JavaScript для Office предоставляет два вспомогательных метода: [Office. Context. Mailbox. convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и [Office. Context. Mailbox. convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

Эти вспомогательные методы выполняют какие-либо действия по разным причинам для следующих двух сценариев, связанных с датами, в расширенном клиенте Outlook, Outlook в Интернете и на мобильных устройствах, что позволяет поднимать "однократная однократная" для разных клиентов вашей надстройки.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Сценарий A. Отображение времени создания или изменения элементов

При отображении времени создания (**Item.dateTimeCreated**) или времени изменения (**Item.dateTimeModified**) элемента в пользовательском интерфейсе метод **convertToLocalClientTime** используется в первый раз для преобразования объекта **Date**, предоставленного этими свойствами, с целью получения представления словаря в соответствующем местном времени. Затем отображаются части даты словаря. Ниже приведен пример этого сценария.


```js
// This date is UTC-correct.
var myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
var myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Обратите внимание, что **convertToLocalClientTime** следит за различием между богатым клиентом Outlook и Outlook в Интернете или мобильных устройствах:


- Если метод **convertToLocalClientTime** воспринимает текущий узел как расширенный клиент, метод преобразует представление **Date** в представление словаря с использованием местного часового пояса клиентского компьютера, что согласуется с остальным пользовательским интерфейсом расширенного клиента.
    
- Если **convertToLocalClientTime** обнаруживает, что текущий узел находится в Outlook в Интернете или на мобильных устройствах, метод преобразует представление **даты** с указанием в формате UTC в формат словаря в часовом поясе центра администрирования Exchange в соответствии с остальной частью Outlook в веб-интерфейсе или в пользовательском интерфейсе для мобильных устройств.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Сценарий Б. Отображение дат начала и окончания в форме создания встречи

При получении в качестве ввода разных частей значения времени и даты, представленных в формате местного времени, и предоставлении этого ввода значения словаря как времени начала или окончания в форме встречи, сначала используйте вспомогательный метод **convertToUtcClientTime**, чтобы преобразовать значение словаря в объект **Date**, соответствующий формату UTC.

В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения берут за основу местное время, зависящее от ведущего приложения.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, соответствующие формату UTC. Затем передайте эти объекты **Date** как аргументы для параметров _Start_ и _End_ метода **Mailbox.displayNewAppointmentForm**, чтобы отобразить форму новой встречи.

Обратите внимание, что **convertToUtcClientTime** следит за различием между богатым клиентом Outlook и Outlook в Интернете или мобильных устройствах:


- Если метод **convertToUtcClientTime** обнаруживает, что текущий узел является расширенным клиентом Outlook, метод просто преобразует представление словаря в объект **Date**. Этот объект **Date** соответствует формату UTC, что и ожидается в методе **displayNewAppointmentForm**.
    
- Если **convertToUtcClientTime** обнаруживает, что текущий узел находится в Outlook в Интернете или на мобильных устройствах, метод преобразует формат словаря значений даты и времени, выраженный в часовом поясе центра администрирования Exchange, в объект **Date** . Этот объект **Date** соответствует формату UTC, что и ожидается в методе **displayNewAppointmentForm**.
    

## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)
    


