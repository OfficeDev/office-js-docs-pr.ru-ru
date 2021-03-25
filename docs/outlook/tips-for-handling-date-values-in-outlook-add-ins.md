---
title: Обработка значений дат в надстройках Outlook
description: API JavaScript Office использует объект Дата JavaScript для большей части хранения и ирисовки дат и времени.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 4f0e93d147760f91c55fd5666f02b2c6cc5d5470
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/24/2021
ms.locfileid: "51177994"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Советы по использованию значений дат в надстройках Outlook

API JavaScript Office использует объект Дата [JavaScript](https://www.w3schools.com/jsref/jsref_obj_date.asp) для большей части хранения и ирисовки дат и времени. 

Этот объект предоставляет такие методы, как `Date` [getUTCDate,](https://www.w3schools.com/jsref/jsref_getutcdate.asp) [getUTCHour,](https://www.w3schools.com/jsref/jsref_getutchours.asp) [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)и [toUTCString,](https://www.w3schools.com/jsref/jsref_toutcstring.asp)которые возвращают запрашиваемую дату или значение времени в соответствии со временем универсального скоординированного времени (UTC).

Объект также предоставляет другие методы, такие как `Date` [getDate,](https://www.w3schools.com/jsref/jsref_getutcdate.asp) [getHour,](https://www.w3schools.com/jsref/jsref_getutchours.asp) [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)и [toString,](https://www.w3schools.com/jsref/jsref_tostring_date.asp)которые возвращают запрашиваемую дату или время в соответствии с "местным временем".

Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, работающих на клиентском компьютере с Windows, вызов JavaScript возвращает дату в зависимости от часового пояса, установленного в Windows на `getDate` клиентском компьютере.

В следующем примере создается объект в локальное время и вызывается преобразование этой даты в строку даты `Date` `myLocalDate` в `toUTCString` UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Хотя объект JavaScript можно использовать для получения значения даты или времени на основе часового пояса UTC или часового пояса клиентского компьютера, объект Date ограничен в одном отношении — он не предоставляет методы возврата даты или значения времени для любого другого определенного часового `Date` пояса.  Например, если клиентский компьютер установлен для восточного стандартного времени (EST), не существует метода, который позволяет получить значение часа, за исключением EST или UTC, таких как Тихоокеанское стандартное время `Date` (PST).


## <a name="date-related-features-for-outlook-add-ins"></a>Функции надстроек Outlook, связанные с датой

Вышеупомянутое ограничение JavaScript имеет значение для вас при использовании API JavaScript Office для обработки значений даты или времени в надстройки Outlook, которые работают в богатом клиенте Outlook, а также в Outlook на веб-или мобильных устройствах.


### <a name="time-zones-for-outlook-clients"></a>Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.

|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. Большинство браузеров используют часовой пояс клиентского компьютера для отображения значений даты или времени объекта `Date` JavaScript.<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. На Mac, если пользователь меняет часовой пояс на клиентский компьютер, Outlook на Mac также будет побуждать пользователя обновлять часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователь задает это значение часового пояса (и предпочтительный язык) при первом входе в Outlook на веб-или мобильных устройствах. <br/><br/>Outlook на веб-и мобильных устройствах использует этот часовой пояс для отображения значений даты или времени в пользовательском интерфейсе.|

Так как богатый клиент Outlook использует часовой пояс клиентского компьютера, а пользовательский интерфейс Outlook на веб-и мобильных устройствах использует часовой пояс EAC, местное время для той же надстройки, установленной для одного и того же почтового ящика, может быть разным при работе в клиенте Outlook и в Outlook на веб-или мобильных устройствах. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.


### <a name="date-related-api"></a>Интерфейс API, связанный с датой

Ниже представлены свойства и методы API JavaScript Office, которые поддерживают функции, связанные с датами.

|Элемент API|Представление часового пояса|Пример в расширенном клиенте Outlook|Пример в Outlook на веб-или мобильных устройствах|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#timezone)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. В Outlook на веб-и мобильных устройствах это свойство возвращает часовой пояс EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Каждое из этих свойств возвращает объект `Date` JavaScript. Это значение UTC-правильно, как показано в следующем примере, имеет такое же значение в богатом клиенте `Date` Outlook, Outlook на веб-и `myUTCDate` мобильных устройствах.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Однако вызов возвращает значение даты в часовом поясе клиентского компьютера, соответствующее часовому поясу, используемому для отображения значений времени даты в клиентской интерфейсе Outlook, но может быть иным, чем часовой пояс EAC, который Outlook на веб-и мобильных устройствах использует в своем пользовательском  `myDate.getDate` интерфейсе.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.|
|[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Для каждого из _параметров_ _"Начните" и "Конец"_ требуется объект JavaScript. `Date` Аргументы должны быть корректно UTC независимо от часового пояса, используемого в пользовательском интерфейсе богатого клиента Outlook или Outlook на веб-или мобильных устройствах.|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Вспомогательные методы для сценариев, связанных с датами


Как описано в предыдущих разделах, так как "локальное время" для пользователя в Outlook на веб-или мобильных устройствах может быть разным для богатого клиента Outlook, но объект даты **JavaScript** поддерживает преобразование только в часовой пояс клиента или UTC, API JavaScript Office предоставляет два метода помощника: [Office.context.mailbox.convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и [Office.context.mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

Эти методы помощника заботятся о необходимости по-разному обрабатывать дату или время для следующих двух сценариев, связанных с датой, в клиенте Outlook с богатыми данными, Outlook на веб-устройствах и мобильных устройствах, тем самым укрепляя "write-once" для разных клиентов надстройки.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Сценарий A. Отображение времени создания или изменения элементов

Если отображается время создания элемента () или время изменения (в пользовательском интерфейсе, сначала используйте для преобразования объекта, предоставленного этими свойствами, чтобы получить представление словаря в соответствующее `Item.dateTimeCreated` `Item.dateTimeModified` `convertToLocalClientTime` `Date` локальное время. Затем отображаются части даты словаря. Ниже приведен пример этого сценария.


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

Обратите внимание, что учитывает разницу между клиентом с богатыми клиентами Outlook и `convertToLocalClientTime` Outlook на веб-или мобильных устройствах:


- Если обнаруживает, что текущее приложение является богатым клиентом, метод преобразует представление в представление словаря в том же часовом поясе клиентского компьютера, в соответствии с остальной частью богатого пользовательского интерфейса `convertToLocalClientTime` `Date` клиента.
    
- Если обнаруживает текущее приложение Outlook на веб-или мобильных устройствах, метод преобразует представление UTC-correct в формат словаря часового пояса EAC, соответствующее остальной части Outlook на пользовательском интерфейсе веб-или мобильных `convertToLocalClientTime` `Date` устройств.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Сценарий Б. Отображение дат начала и окончания в форме создания встречи

Если вы получаете в качестве ввода различные части значения даты, представленного в локальное время, и хотели бы предоставить это значение словаря в виде начала или окончания в форме встречи, сначала используйте метод помощника для преобразования значения словаря в объект, правильный `convertToUtcClientTime` `Date` UTC.

В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения основаны на локальном времени, в зависимости от клиентской платформы.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, правильные относительно UTC. Затем передайте эти объекты в качестве аргументов для параметров Start и End метода для отображения `Date` новой формы   `Mailbox.displayNewAppointmentForm` встречи.

Обратите внимание, что учитывает разницу между клиентом с богатыми клиентами Outlook и `convertToUtcClientTime` Outlook на веб-или мобильных устройствах:


- Если обнаруживает, что текущее приложение является богатым клиентом Outlook, метод просто преобразует представление словаря `convertToUtcClientTime` в `Date` объект. Этот `Date` объект является корректным по UTC, как и ожидалось `displayNewAppointmentForm` .
    
- Если обнаруживается текущее приложение Outlook на веб-или мобильных устройствах, метод преобразует формат словаря значений даты и времени, выраженных в часовом поясе `convertToUtcClientTime` EAC, в `Date` объект. Этот `Date` объект является корректным по UTC, как и ожидалось `displayNewAppointmentForm` .
    
## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)
