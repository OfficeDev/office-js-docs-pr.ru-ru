---
title: Обработка значений дат в надстройках Outlook
description: API Office JavaScript использует объект Дата JavaScript для большей части хранения и ирисовки дат и времени.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 46be9e7e3c952d08addcf8ef761a259f8c0d1d84c1bc3b0bb61cbb40c07ce35b
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093324"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Советы по использованию значений дат в надстройках Outlook

API Office JavaScript использует объект Дата [JavaScript](https://www.w3schools.com/jsref/jsref_obj_date.asp) для большей части хранения и ирисовки дат и времени. 

Этот объект предоставляет такие методы, как `Date` [getUTCDate,](https://www.w3schools.com/jsref/jsref_getutcdate.asp) [getUTCHour,](https://www.w3schools.com/jsref/jsref_getutchours.asp) [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp)и [toUTCString,](https://www.w3schools.com/jsref/jsref_toutcstring.asp)которые возвращают запрашиваемую дату или значение времени в соответствии со временем универсального скоординированного времени (UTC).

Объект также предоставляет другие методы, такие как `Date` [getDate,](https://www.w3schools.com/jsref/jsref_getutcdate.asp) [getHour,](https://www.w3schools.com/jsref/jsref_getutchours.asp) [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp)и [toString,](https://www.w3schools.com/jsref/jsref_tostring_date.asp)которые возвращают запрашиваемую дату или время в соответствии с "местным временем".

Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, работающих Windows клиентском компьютере, вызов JavaScript возвращает дату в зависимости от часового пояса, установленного Windows на клиентском `getDate` компьютере.

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

Вышеупомянутое ограничение JavaScript имеет значение для вас, когда вы используете API Office JavaScript для обработки значений даты или времени в надстройки Outlook, которые работают в Outlook клиенте, а также в Outlook в Интернете или мобильных устройствах.


### <a name="time-zones-for-outlook-clients"></a>Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.

|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. Большинство браузеров используют часовой пояс клиентского компьютера для отображения значений даты или времени объекта `Date` JavaScript.<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. На Mac, если пользователь изменяет часовой пояс на клиентский компьютер, Outlook mac будет побуждать пользователя обновить часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователь задает это значение часового пояса (и предпочтительный язык), когда он в первый раз Outlook в Интернете или мобильные устройства. <br/><br/>Outlook в Интернете и мобильные устройства используют этот часовой пояс для отображения значений даты или времени в пользовательском интерфейсе.|

Так как Outlook клиент использует часовой пояс клиентского компьютера, а пользовательский интерфейс Outlook в Интернете и мобильных устройств использует часовой пояс EAC, местное время для одной и той же надстройки, установленной для одного и того же почтового ящика, может быть разным при работе в Outlook клиенте и на Outlook в Интернете или мобильных устройствах. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.


### <a name="date-related-api"></a>Интерфейс API, связанный с датой

Ниже представлены свойства и методы, Office API JavaScript, поддерживают функции, связанные с датами.

|Элемент API|Представление часового пояса|Пример в расширенном клиенте Outlook|Пример в Outlook в Интернете или мобильных устройствах|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#timeZone)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. В Outlook в Интернете и мобильных устройствах это свойство возвращает часовой пояс EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) и [Office.context.mailbox.item.dateTimeModified](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties)|Каждое из этих свойств возвращает объект `Date` JavaScript. Это значение правильно UTC, как показано в следующем примере, имеет одинаковое значение в Outlook `Date` `myUTCDate` клиенте, Outlook в Интернете и мобильных устройствах.<br/><br/>`var myDate = Office.mailbox.item.dateTimeCreated;`<br/>`var myUTCDate = myDate.getUTCDate;`<br/><br/>Однако вызов возвращает значение даты в часовом поясе клиентского компьютера, которое соответствует часовому поясу, используемому для отображения значений времени даты в интерфейсе клиента Outlook, но может быть иным, чем часовой пояс EAC, используемый Outlook в Интернете и мобильными устройствами в пользовательском `myDate.getDate` интерфейсе.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если элемент создан в 9 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11 часов утра в формате времени UTC, для метода<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.|
|[Office.context.mailbox.displayNewAppointmentForm](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)|Для каждого из _параметров_ _"Начните" и "Конец"_ требуется объект JavaScript. `Date` Аргументы должны быть корректно UTC независимо от часового пояса, используемого в пользовательском интерфейсе богатого клиента Outlook или Outlook в Интернете или мобильных устройств.|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|Если значениями начального и конечного времени для формы встречи являются 9 и 11 часов утра в формате времени UTC, следует убедиться, что аргументы `start` и `end` правильны относительно формата времени UTC. Это означает, что для метода<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Вспомогательные методы для сценариев, связанных с датами


Как описано в предыдущих разделах, так как "локальное время" для пользователя в Outlook в Интернете или мобильных устройствах может быть разным для богатого клиента Outlook, но объект даты **JavaScript** поддерживает преобразование только в часовой пояс клиента или UTC, API javaScript Office предоставляет два метода: [Office.context.mailbox.convertToLocalClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) и [Office.context.mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).

Эти методы помощника заботятся о необходимости по-разному обрабатывать дату или время для следующих двух сценариев, связанных с датами, в Outlook клиенте, Outlook в Интернете и мобильных устройствах, тем самым усиливая "один раз записи" для разных клиентов надстройки.


### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Сценарий A. Отображение времени создания или изменения элементов

Если отображается время создания элемента () или время изменения (в пользовательском интерфейсе, сначала используйте для преобразования объекта, предоставленного этими свойствами, чтобы получить представление словаря в соответствующее `Item.dateTimeCreated` `Item.dateTimeModified` `convertToLocalClientTime` `Date` локальное время. Затем отображаются части даты словаря. Ниже приводится пример этого сценария.


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

Обратите внимание на разницу между богатым клиентом Outlook клиентом и Outlook в Интернете `convertToLocalClientTime` или мобильными устройствами:


- Если обнаруживает, что текущее приложение является богатым клиентом, метод преобразует представление в представление словаря в том же часовом поясе клиентского компьютера, в соответствии с остальной частью богатого пользовательского интерфейса `convertToLocalClientTime` `Date` клиента.
    
- Если обнаруживается текущее приложение Outlook в Интернете или мобильных устройств, метод преобразует представление UTC-правильно в формат словаря в часовом поясе EAC, в соответствии с остальной частью пользовательского интерфейса Outlook в Интернете или мобильных `convertToLocalClientTime` `Date` устройств.
    

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Сценарий Б. Отображение дат начала и окончания в форме создания встречи

Если вы получаете в качестве ввода различные части значения даты, представленного в локальное время, и хотели бы предоставить это значение словаря в виде начала или окончания в форме встречи, сначала используйте метод помощника для преобразования значения словаря в объект, правильный `convertToUtcClientTime` `Date` UTC.

В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения основаны на локальном времени, в зависимости от клиентской платформы.

```js
var myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
var myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, правильные относительно UTC. Затем передайте эти объекты в качестве аргументов для параметров Start и End метода для отображения `Date` новой формы   `Mailbox.displayNewAppointmentForm` встречи.

Обратите внимание на разницу между богатым клиентом Outlook клиентом и Outlook в Интернете `convertToUtcClientTime` или мобильными устройствами:


- Если обнаруживает, что текущее приложение является Outlook клиентом, метод просто преобразует представление словаря `convertToUtcClientTime` в `Date` объект. Этот `Date` объект является корректным по UTC, как и ожидалось `displayNewAppointmentForm` .
    
- Если обнаруживается текущее приложение Outlook в Интернете или мобильных устройств, метод преобразует формат словаря значений даты и времени, выраженных в часовом поясе `convertToUtcClientTime` EAC, в `Date` объект. Этот `Date` объект является корректным по UTC, как и ожидалось `displayNewAppointmentForm` .
    
## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)
