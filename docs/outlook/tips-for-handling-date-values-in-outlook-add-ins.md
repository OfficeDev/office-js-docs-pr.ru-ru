---
title: Обработка значений дат в надстройках Outlook
description: API JavaScript для Office использует объект JavaScript Date для большей части хранения и извлечения даты и времени.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 3645d3f91b07c847e05a45563f75c5fc0cbe0135
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611640"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Советы по использованию значений дат в надстройках Outlook

API JavaScript для Office использует объект JavaScript [Date](https://www.w3schools.com/jsref/jsref_obj_date.asp) для большей части хранения и извлечения даты и времени. 

Этот `Date` объект предоставляет методы, такие как [жетуткдате](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [Жетутчаур](https://www.w3schools.com/jsref/jsref_getutchours.asp), [жетуткминутес](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)и [вызовом toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), которые возвращают запрошенное значение даты или времени в соответствии с всеобщим скоординированным временем (UTC).

`Date`Объект также предоставляет другие методы, например GETDATE, [GETDATE](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp)GETDATE и [ToString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), которые возвращают запрошенную дату или время в соответствии с "местным временем". [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)

Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, запущенных на клиентском компьютере под управлением Windows, вызов JavaScript `getDate` возвращает дату в соответствии с часовым поясом, установленным в Windows на клиентском компьютере.

В следующем примере создается `Date` объект `myLocalDate` по местному времени, а затем выполняется вызов `toUTCString` для преобразования даты в строку даты в формате UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Несмотря на то, что вы можете использовать `Date` объект JavaScript для получения значения даты или времени на основе UTC или часового пояса клиентского компьютера, объект **Date** ограничен в одном отношениях не предоставляет методы для возврата значения даты или времени для любого другого определенного часового пояса. Например, если ваш клиентский компьютер настроен на зимнее стандартное время (EST), не существует `Date` метода, позволяющего получить значение часа, отличное от в средстве EST или UTC, например по тихоокеанскому времени (PST).


## <a name="date-related-features-for-outlook-add-ins"></a>Функции надстроек Outlook, связанные с датой

При использовании API JavaScript для Office для обработки значений даты и времени в надстройках Outlook, выполняемых в расширенном клиенте Outlook, а также в Outlook в Интернете или на мобильных устройствах, упомянутые выше ограничения JavaScript имеют недостаточное значение.


### <a name="time-zones-for-outlook-clients"></a>Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.

|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. Большинство браузеров используют часовой пояс клиентского компьютера для отображения значений даты и времени `Date` объекта JavaScript.<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. Если пользователь изменяет часовой пояс на клиентском компьютере в Mac, Outlook в MAC-адресе будет предлагать пользователю обновить часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователь задает это значение часового пояса (и предпочитаемый язык), когда он впервые выполняет вход в Outlook в Интернете или на мобильных устройствах. <br/><br/>В Outlook в Интернете и на мобильных устройствах этот часовой пояс используется для отображения значений даты и времени в пользовательском интерфейсе.|

Так как расширенный клиент Outlook использует часовой пояс клиентского компьютера, а пользовательский интерфейс Outlook в Интернете и на мобильных устройствах использует часовой пояс центра администрирования Exchange, местное время для той же надстройки, установленной для того же почтового ящика, может отличаться при работе в расширенном клиенте Outlook и в Outlook в Интернете или на мобильных устройствах. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.


### <a name="date-related-api"></a>Интерфейс API, связанный с датой

Ниже приведены свойства и методы API JavaScript для Office, которые поддерживают функции, связанные с датами.

**Элемент API**|**Представление часового пояса**|**Пример в расширенном клиенте Outlook**|**Пример в Outlook в Интернете или на мобильных устройствах**
--------------|----------------------------|-------------------------------------|-------------------
[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview#timezone)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. В Outlook в Интернете и мобильных устройствах это свойство возвращает часовой пояс центра администрирования Exchange. |EST|PST
[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Каждое из этих свойств возвращает объект JavaScript `Date` . Это `Date` значение указано в формате UTC, как показано в следующем примере — `myUTCDate` имеет то же значение, что и в расширенном клиенте Outlook, Outlook в Интернете и на мобильных устройствах.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Однако вызов `myDate.getDate` возвращает значение даты в часовом поясе клиентского компьютера, которое согласуется с часовым поясом, используемым для отображения значений даты и времени в пользовательском интерфейсе Outlook с расширенными возможностями, но может отличаться от часовых поясов, которые Outlook в Интернете и мобильные устройства используют в своем пользовательском интерфейсе.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.
[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Для каждого из параметров _Start_ и _End_ требуется объект JavaScript `Date` . Аргументы должны быть правильно заданы в формате UTC независимо от часового пояса, используемого в пользовательском интерфейсе в расширенном клиенте Outlook или в Интернете или на мобильных устройствах.|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>

## <a name="helper-methods-for-date-related-scenarios"></a>Вспомогательные методы для сценариев, связанных с датами


Как описано в предыдущих разделах, так как "Местное время" для пользователя в Outlook в Интернете или мобильных устройствах может различаться в расширенном клиенте Outlook, но объект JavaScript **Date** поддерживает преобразование только в часовой пояс клиентского компьютера или в формате UTC, API JavaScript для Office предоставляет два вспомогательных метода: [Office. Context. Mailbox. convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и [Office. Context. Mailbox. convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

Эти вспомогательные методы выполняют какие-либо действия по разным причинам для следующих двух сценариев, связанных с датами, в расширенном клиенте Outlook, Outlook в Интернете и на мобильных устройствах, что позволяет поднимать "однократная однократная" для разных клиентов вашей надстройки.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Сценарий A. Отображение времени создания или изменения элементов

При отображении времени создания элемента ( `Item.dateTimeCreated` ) или времени изменения ( `Item.dateTimeModified` в пользовательском интерфейсе сначала используется `convertToLocalClientTime` для преобразования `Date` объекта, предоставленного этими свойствами, для получения представления словаря в соответствующее местное время. Затем отображаются части даты словаря. Ниже приведен пример этого сценария.


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

Обратите внимание, что в Outlook `convertToLocalClientTime` в Интернете или на мобильных устройствах применяется разница между расширенным клиентом Outlook и Outlook в Интернете.


- Если `convertToLocalClientTime` обнаруживается, что текущий узел является полнофункциональным клиентом, метод преобразует `Date` представление в словарь в том же часовом поясе клиентского компьютера, который согласуется с остальным пользовательским интерфейсом расширенного клиента.
    
- Если `convertToLocalClientTime` обнаруживается, что текущий узел находится в Outlook в Интернете или на мобильных устройствах, метод преобразует представление правильного формата времени в формате UTC `Date` в формат словаря в часовом поясе центра администрирования Exchange в соответствии с остальной частью Outlook в веб-интерфейсе или пользовательском интерфейсе мобильных устройств.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Сценарий Б. Отображение дат начала и окончания в форме создания встречи

Если вы используете в качестве входных данных различные части значения даты и времени, представленные в местном времени, и хотите предоставить это входное значение словаря в качестве времени начала или окончания в форме встречи, сначала используйте `convertToUtcClientTime` вспомогательный метод для преобразования значения словаря в соответствующий объект в формате UTC `Date` .

В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения берут за основу местное время, зависящее от ведущего приложения.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, соответствующие формату UTC. Затем передайте эти `Date` объекты в качестве аргументов для параметров _Start_ и _End_ `Mailbox.displayNewAppointmentForm` метода, чтобы отобразить форму новой встречи.

Обратите внимание, что в Outlook `convertToUtcClientTime` в Интернете или на мобильных устройствах применяется разница между расширенным клиентом Outlook и Outlook в Интернете.


- Если `convertToUtcClientTime` обнаруживается, что текущий узел является расширенным клиентом Outlook, метод просто преобразует представление словаря в `Date` объект. Этот `Date` объект является правильным временем в формате UTC, как ожидалось `displayNewAppointmentForm` .
    
- Если `convertToUtcClientTime` обнаруживается, что текущий узел находится в Outlook в Интернете или на мобильных устройствах, метод преобразует формат словаря значений даты и времени, выраженный в часовом поясе центра администрирования Exchange, в `Date` объект. Этот `Date` объект является правильным временем в формате UTC, как ожидалось `displayNewAppointmentForm` .
    
## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)