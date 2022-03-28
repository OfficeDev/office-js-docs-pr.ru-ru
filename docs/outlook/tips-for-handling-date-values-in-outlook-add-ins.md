---
title: Обработка значений дат в надстройках Outlook
description: API Office JavaScript использует объект Дата JavaScript для большей части хранения и ирисовки дат и времени.
ms.date: 10/31/2019
ms.localizationpriority: medium
ms.openlocfilehash: adcf7cebd93a5881094a843d19fd65f95ae459a3
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64484486"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Советы по использованию значений дат в надстройках Outlook

API Office JavaScript использует объект Дата [JavaScript для](https://www.w3schools.com/jsref/jsref_obj_date.asp) большей части хранения и ирисовки дат и времени. 

`Date` Этот объект предоставляет такие методы, как [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) и [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), которые возвращают запрашиваемую дату или значение времени в соответствии со временем универсального скоординированного времени (UTC).

Объект `Date` также предоставляет другие методы, такие как [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) и [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), которые возвращают запрашиваемую дату или время в соответствии с "местным временем".

Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, работающих Windows клиентском компьютере, вызов JavaScript `getDate`возвращает дату в зависимости от часового пояса, установленного Windows на клиентском компьютере.

В следующем примере создается `Date` объект в `myLocalDate` локальное время и `toUTCString` вызывается преобразование этой даты в строку даты в UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
var myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Хотя объект JavaScript `Date` можно использовать для получения значения даты или времени на основе часового пояса UTC или часового пояса клиентского компьютера, объект **Date** ограничен в одном отношении — он не предоставляет методы возврата даты или значения времени для любого другого определенного часового пояса. Например, если клиентский компьютер установлен для восточного стандартного времени (EST), `Date` не существует метода, который позволяет получить значение часа, за исключением EST или UTC, таких как Тихоокеанское стандартное время (PST).


## <a name="date-related-features-for-outlook-add-ins"></a>Функции надстроек Outlook, связанные с датой

Вышеупомянутое ограничение JavaScript имеет значение для вас, когда вы используете API Office JavaScript для обработки значений даты или времени в надстройки Outlook, которые работают в Outlook клиенте, а также в Outlook в Интернете или мобильных устройствах.


### <a name="time-zones-for-outlook-clients"></a>Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.

|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. Большинство браузеров используют часовой пояс клиентского компьютера для отображения значений даты или времени объекта JavaScript `Date` .<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. На Mac, если пользователь изменяет часовой пояс на клиентский компьютер, Outlook mac будет побуждать пользователя обновить часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователь задает это значение часового пояса (и предпочтительный язык), когда он в первый раз Outlook в Интернете или мобильные устройства. <br/><br/>Outlook в Интернете и мобильные устройства используют этот часовой пояс для отображения значений даты или времени в пользовательском интерфейсе.|

Так как Outlook клиент использует часовой пояс клиентского компьютера, а пользовательский интерфейс Outlook в Интернете и мобильных устройств использует часовой пояс EAC, местное время для одной и той же надстройки, установленной для одного и того же почтового ящика, может быть разным при работе в Outlook клиенте и Outlook в Интернете или мобильных устройствах. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.


### <a name="date-related-api"></a>Интерфейс API, связанный с датой

Ниже представлены свойства и методы в API Office JavaScript, поддерживают функции, связанные с датами.

|Элемент API|Представление часового пояса|Пример в расширенном клиенте Outlook|Пример в Outlook в Интернете или мобильных устройствах|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. В Outlook в Интернете и мобильных устройствах это свойство возвращает часовой пояс EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) и [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Каждое из этих свойств возвращает объект JavaScript `Date` . Это `Date` значение UTC-правильно, как показано в следующем примере, `myUTCDate` имеет одинаковое значение в Outlook клиенте, Outlook в Интернете и мобильных устройствах.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>`myDate.getDate` Однако вызов возвращает значение даты в часовом поясе клиентского компьютера, которое соответствует часовому поясу, используемому для отображения значений времени даты в богатом клиентом интерфейсе Outlook, но может быть иным, чем часовой пояс EAC, который Outlook в Интернете и мобильных устройств, используемых в пользовательском интерфейсе.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|Для каждого из параметров  _"_ Начните" и " _Конец_ " требуется объект JavaScript `Date` . Аргументы должны быть корректно UTC независимо от часового пояса, используемого в пользовательском интерфейсе Outlook клиента или Outlook в Интернете или мобильных устройств.|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Вспомогательные методы для сценариев, связанных с датами


Как описано в предыдущих разделах, так как "местное время" для пользователя в Outlook в Интернете или мобильных устройствах может быть разным для богатого клиента Outlook, но объект JavaScript **Date** поддерживает преобразование только в часовой пояс клиента или UTC, API javaScript Office предоставляет два дополнительных метода: [Office .context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) [и Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

Эти методы помощника заботятся о необходимости по-разному обрабатывать дату или время для следующих двух сценариев, связанных с датами, в Outlook клиенте, Outlook в Интернете и мобильных устройствах, тем самым усиливая "write-once" для разных клиентов надстройки.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Сценарий A. Отображение времени создания или изменения элементов

Если отображается время создания элемента () или время изменения (`Item.dateTimeCreated``Item.dateTimeModified`в пользовательском интерфейсе, `convertToLocalClientTime` `Date` сначала используйте для преобразования объекта, предоставленного этими свойствами, чтобы получить представление словаря в соответствующее локальное время. Затем отображаются части даты словаря. Ниже приводится пример этого сценария.


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

Обратите внимание `convertToLocalClientTime` на разницу между богатым клиентом Outlook клиентом и Outlook в Интернете или мобильными устройствами:


- Если `convertToLocalClientTime` обнаруживает, что текущее приложение является богатым клиентом, `Date` метод преобразует представление в представление словаря в том же часовом поясе клиентского компьютера, в соответствии с остальной частью богатого пользовательского интерфейса клиента.
    
- `convertToLocalClientTime` Если обнаруживается текущее приложение Outlook в Интернете или мобильных устройств, метод преобразует представление UTC-правильно `Date` в формат словаря часового пояса EAC, в соответствии с остальной частью пользовательского интерфейса Outlook в Интернете или мобильных устройств.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Сценарий Б. Отображение дат начала и окончания в форме создания встречи

Если вы получаете в качестве ввода различные части значения даты, представленного в локальное время, и хотели бы предоставить это значение словаря в виде начала или окончания в форме встречи, `convertToUtcClientTime` сначала используйте метод помощника для преобразования значения словаря в объект, правильный UTC `Date` .

В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения основаны на локальном времени, в зависимости от клиентской платформы.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, правильные относительно UTC. Затем передайте `Date` эти объекты в качестве аргументов для параметров _Start_ и _End_ `Mailbox.displayNewAppointmentForm` метода для отображения новой формы встречи.

Обратите внимание `convertToUtcClientTime` на разницу между богатым клиентом Outlook клиентом и Outlook в Интернете или мобильными устройствами:


- Если `convertToUtcClientTime` обнаруживает, что текущее приложение является Outlook клиентом, метод просто преобразует представление словаря в `Date` объект. Этот `Date` объект является корректным по UTC, как и ожидалось `displayNewAppointmentForm`.
    
- Если `convertToUtcClientTime` обнаруживается текущее приложение Outlook в Интернете или мобильных устройств, метод преобразует формат словаря значений даты и времени, выраженных в часовом поясе EAC`Date`, в объект. Этот `Date` объект является корректным по UTC, как и ожидалось `displayNewAppointmentForm`.
    
## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)
