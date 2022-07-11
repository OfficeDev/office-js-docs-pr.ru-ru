---
title: Обработка значений дат в надстройках Outlook
description: API JavaScript для Office использует объект даты JavaScript для большей части хранилища и извлечения дат и времени.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 49de8db712400e006dc919e9ad62ae6cbaaa11cf
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713079"
---
# <a name="tips-for-handling-date-values-in-outlook-add-ins"></a>Советы по использованию значений дат в надстройках Outlook

API JavaScript для Office использует объект [даты](https://www.w3schools.com/jsref/jsref_obj_date.asp) JavaScript для большей части хранилища и извлечения дат и времени.

`Date` Этот объект предоставляет такие методы, как [getUTCDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getUTCHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getUTCMinutes](https://www.w3schools.com/jsref/jsref_getutcminutes.asp) и [toUTCString](https://www.w3schools.com/jsref/jsref_toutcstring.asp), которые возвращают запрошенное значение даты или времени в соответствии со временем UTC.

Объект `Date` также предоставляет другие методы, такие как [getDate](https://www.w3schools.com/jsref/jsref_getutcdate.asp), [getHour](https://www.w3schools.com/jsref/jsref_getutchours.asp), [getMinutes](https://www.w3schools.com/jsref/jsref_getminutes.asp) и [toString](https://www.w3schools.com/jsref/jsref_tostring_date.asp), которые возвращают запрашиваемую дату или время в соответствии с "локальным временем".

Понятие "местного времени" в значительной мере определяется браузером и операционной системой на клиентском компьютере. Например, в большинстве браузеров, работающих на клиентском компьютере под управлением Windows, вызов JavaScript `getDate`возвращает дату на основе часового пояса, установленного в Windows на клиентском компьютере.

В следующем примере создается объект `Date` в `myLocalDate` местное время и `toUTCString` вызывается преобразование этой даты в строку даты в формате UTC.

```js
// Create and get the current date represented 
// in the client computer time zone.
const myLocalDate = new Date (); 

// Convert the Date value in the client computer time zone
// to a date string in UTC, and display the string.
document.write ("The current UTC time is " + 
    myLocalDate.toUTCString());
```

Хотя объект JavaScript `Date` можно использовать для получения значения даты или времени на основе времени в формате UTC или часового пояса клиентского компьютера, объект **Date** ограничен в одном отношении. Он не предоставляет методы для возврата значения даты или времени для любого другого определенного часового пояса. Например, если на клиентском компьютере установлено восточное стандартное время (EST), `Date` метод, позволяющий получить значение часа, отличное от значения EST или UTC, например тихоокеанского стандартного времени (PST).

## <a name="date-related-features-for-outlook-add-ins"></a>Функции надстроек Outlook, связанные с датой

Упомянутое ограничение JavaScript влияет на вас, если вы используете API JavaScript для Office для обработки значений даты и времени в надстройки Outlook, которые работают в полнофункциональных клиентах Outlook, а также на Outlook в Интернете или мобильных устройствах.

### <a name="time-zones-for-outlook-clients"></a>Часовые пояса для клиентов Outlook

Во избежание недоразумений дадим определение часовым поясам.

|**Часовой пояс**|**Описание**|
|:-----|:-----|
|Часовой пояс клиентского компьютера|Устанавливается в операционной системе на клиентском компьютере. Большинство браузеров используют часовой пояс клиентского компьютера для отображения значений даты или времени объекта JavaScript `Date` .<br/><br/>В расширенном клиенте Outlook используется этот часовой пояс для отображения значений даты и времени в пользовательском интерфейсе. <br/><br/>Например, на клиентском компьютере под управлением Windows в Outlook используется часовой пояс, установленный в операционной системе Windows в качестве местного часового пояса. Если на компьютере Mac пользователь изменяет часовой пояс на клиентском компьютере, Outlook для Mac также предложит обновить часовой пояс в Outlook.|
|Часовой пояс Центра администрирования Exchange (EAC)|Пользователь задает это значение часового пояса (и предпочитаемый язык) при первом входе Outlook в Интернете или мобильных устройствах. <br/><br/>Outlook в Интернете и мобильные устройства используют этот часовой пояс для отображения значений даты или времени в пользовательском интерфейсе.|

Так как полнофункциональный клиент Outlook использует часовой пояс клиентского компьютера, а пользовательский интерфейс Outlook в Интернете и мобильных устройств использует часовой пояс EAC, местное время для той же надстройки, установленной для того же почтового ящика, может отличаться при работе в полнофункциональных клиентах Outlook и на Outlook в Интернете или мобильных устройствах. Разработчику надстройки Outlook следует продумать ввод и вывод значений даты, чтобы эти значения всегда согласовывались с часовым поясом, который пользователь ожидает увидеть в соответствующем клиенте.

### <a name="date-related-api"></a>Интерфейс API, связанный с датой

Ниже приведены свойства и методы в API JavaScript для Office, которые поддерживают функции, связанные с датами.

|Элемент API|Представление часового пояса|Пример в расширенном клиенте Outlook|Пример на Outlook в Интернете или мобильных устройствах|
|--------------|----------------------------|-------------------------------------|-------------------|
|[Office.context.mailbox.userProfile.timeZone](/javascript/api/outlook/office.userprofile?view=outlook-js-preview&preserve-view=true#outlook-office-userprofile-timezone-member)|В расширенном клиенте Outlook это свойство возвращает часовой пояс клиентского компьютера. На Outlook в Интернете и мобильных устройствах это свойство возвращает часовой пояс EAC. |EST|PST|
|[Office.context.mailbox.item.dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) и [Office.context.mailbox.item.dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|Каждое из этих свойств возвращает объект JavaScript `Date` . Это `Date` значение верно в формате UTC, `myUTCDate` как показано в следующем примере, — имеет то же значение в полнофункциональных клиентах Outlook, Outlook в Интернете и мобильных устройствах.<br/><br/>`const myDate = Office.mailbox.item.dateTimeCreated;`<br/>`const myUTCDate = myDate.getUTCDate;`<br/><br/>Однако вызов возвращает значение даты в часовом поясе клиентского компьютера, которое согласуется с часовой поясом, `myDate.getDate` используемым для отображения значений даты и времени в расширенном клиентском интерфейсе Outlook, но может отличаться от часового пояса EAC, используемого Outlook в Интернете и мобильными устройствами в пользовательском интерфейсе.|Если элемент создается в 9:00 (UTC):<br/><br/>`Office.mailbox.item.`<br/>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11:00 (UTC):<br/><br/>`Office.mailbox.item.`<br/>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.|Если время создания элемента равно 9:00 (UTC):<br/><br/>`Office.mailbox.item.`</br>`dateTimeCreated.getHours` возвращается значение 4 часа утра в формате времени EST.<br/><br/>Если элемент изменен в 11:00 (UTC):<br/><br/>`Office.mailbox.item.`</br>`dateTimeModified.getHours` возвращается значение 6 часов утра в формате времени EST.<br/><br/>Обратите внимание, что если необходимо отобразить время создания или изменения в пользовательском интерфейсе, следует сначала преобразовать время в формат PST, чтобы оно соответствовало формату времени остального пользовательского интерфейса.|
|[Office.context.mailbox.displayNewAppointmentForm](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)|Для каждого _параметра Start_ _и End_ требуется объект JavaScript `Date` . Аргументы должны быть правильными в формате UTC независимо от часового пояса, используемого в пользовательском интерфейсе полнофункциональных клиентов Outlook, Outlook в Интернете или мобильных устройствах.|Если время начала и окончания формы встречи — 9:00 (UTC) и 11:00 (UTC), `start` `end` убедитесь, что аргументы и аргументы указаны в формате UTC. Это означает:<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|Если время начала и окончания формы встречи — 9:00 (UTC) и 11:00 (UTC), `start` `end` убедитесь, что аргументы и аргументы указаны в формате UTC. Это означает:<br/><br/><ul><li>`start.getUTCHours` возвращается значение 9 часов утра в формате времени UTC, а для метода</li><li>`end.getUTCHours` возвращается значение 11 часов утра в формате времени UTC</li></ul>|

## <a name="helper-methods-for-date-related-scenarios"></a>Вспомогательные методы для сценариев, связанных с датами

Как описано в предыдущих разделах, так как "местное время" для пользователя на Outlook в Интернете или мобильных устройствах может отличаться в полнофункциональных клиентах Outlook, но объект даты **JavaScript поддерживает** преобразование только в часовой пояс клиентского компьютера или UTC, API JavaScript для Office предоставляет два вспомогательных метода: [Office.context.mailbox.convertToLocalClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) и [Office.context.mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).

Эти вспомогательные методы по-разному обрабатывают дату или время в следующих двух сценариях, связанных с датой, в полнофункциональных клиентах Outlook Outlook в Интернете и мобильных устройствах, что приводит к "однократной записи" для разных клиентов надстройки.

### <a name="scenario-a-displaying-item-creation-or-modified-time"></a>Сценарий A. Отображение времени создания или изменения элементов

Если отображается время создания элемента (`Item.dateTimeCreated`) или время изменения (`Item.dateTimeModified`в пользовательском интерфейсе, сначала используйте для преобразования объекта, `convertToLocalClientTime` `Date` предоставленного этими свойствами, чтобы получить представление словаря в соответствующее локальное время. Затем отображаются части даты словаря. Ниже приведен пример этого сценария.

```js
// This date is UTC-correct.
const myDate = Office.context.mailbox.item.dateTimeCreated;

// Call helper method to get date in dictionary format, 
// represented in the appropriate local time.
// In an Outlook rich client, this is dictionary format 
// in client computer time zone.
// In Outlook on the web or mobile devices, this dictionary 
// format is in EAC time zone.
const myLocalDictionaryDate = Office.context.mailbox.convertToLocalClientTime(myDate);

// Display different parts of the dictionary date.
document.write ("The item was created at " + myLocalDictionaryDate["hours"] + 
    ":" + myLocalDictionaryDate["minutes"]);)
```

Обратите внимание`convertToLocalClientTime`, что разница между полнофункционалированным клиентом Outlook и Outlook в Интернете или мобильными устройствами:

- Если `convertToLocalClientTime` обнаруживает, что текущее приложение является полнофункционалированным клиентом, `Date` метод преобразует представление в представление словаря в том же часовом поясе клиентского компьютера в соответствии с остальной частью расширенного пользовательского интерфейса клиента.

- `convertToLocalClientTime` Если обнаруживается, что текущее приложение Outlook в Интернете или мобильных устройствах, метод преобразует правильное представление в формате UTC `Date` в формат словаря часового пояса EAC в соответствии с остальной частью пользовательского интерфейса Outlook в Интернете или мобильных устройств.

### <a name="scenario-b-displaying-start-and-end-dates-in-a-new-appointment-form"></a>Сценарий Б. Отображение дат начала и окончания в форме создания встречи

Если вы получаете в качестве входных данных различные части значения даты и времени, представленного в локальном времени, и хотите предоставить это входное значение словаря в качестве времени начала или окончания в форме встречи, `convertToUtcClientTime` сначала используйте вспомогательный метод для преобразования значения словаря в правильный объект UTC `Date` .

В указанном ниже примере предположим, что `myLocalDictionaryStartDate` и `myLocalDictionaryEndDate` — значения даты и времени в формате словаря, полученные от пользователя. Эти значения зависят от местного времени и зависят от клиентской платформы.

```js
const myUTCCorrectStartDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryStartDate);
const myUTCCorrectEndDate = Office.context.mailbox.convertToUtcClientTime(myLocalDictionaryEndDate);

```

В результате получаются значения `myUTCCorrectStartDate` и `myUTCCorrectEndDate`, правильные относительно UTC. Затем передайте эти `Date` объекты в качестве аргументов для параметров _start_ и _End_ `Mailbox.displayNewAppointmentForm` метода, чтобы отобразить новую форму встречи.

Обратите внимание`convertToUtcClientTime`, что разница между полнофункционалированным клиентом Outlook и Outlook в Интернете или мобильными устройствами:

- Если `convertToUtcClientTime` обнаруживает, что текущее приложение является полнофункционалированным клиентом Outlook, метод просто преобразует представление словаря в `Date` объект. Этот `Date` объект является правильным в формате UTC, как и ожидалось `displayNewAppointmentForm`.

- Если `convertToUtcClientTime` обнаруживает, что текущее приложение Outlook в Интернете или мобильных устройствах, метод преобразует формат словаря значений даты и времени, выраженных в часовом поясе EAC`Date`, в объект. Этот `Date` объект является правильным в формате UTC, как и ожидалось `displayNewAppointmentForm`.

## <a name="see-also"></a>См. также

- [Развертывание и установка надстроек Outlook для тестирования](testing-and-tips.md)
