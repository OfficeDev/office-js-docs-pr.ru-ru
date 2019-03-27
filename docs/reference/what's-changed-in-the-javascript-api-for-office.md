# <a name="whats-changed-in-the-javascript-api-for-office"></a>Изменения API JavaScript для Office

В интерфейс API JavaScript для Office периодически добавляются новые и обновленные объекты, методы, свойства, события и перечисления для расширения возможностей ваших Надстройки Office. Используйте следующие ссылки, чтобы ознакомиться с новыми и обновленными элементами API.

Для разработки надстроек с использованием новых элементов API вам потребуется [обновить файлы API JavaScript для Office в проекте](/office/dev/add-ins/develop/update-your-javascript-api-for-office-and-manifest-schema-version).

Сведения обо всех элементах API, в том числе о тех, которые не изменились по сравнению с предыдущими версиями, см. в статье [API JavaScript для Office](javascript-api-for-office.md).

## <a name="new-and-updated-apis"></a>Новые и обновленные интерфейсы API

### <a name="new-and-updated-objects"></a>Новые и обновленные объекты

|**Object**|**Описание**|**Добавленная или обновленная версия **|
|:-----|:-----|:-----|
|`Item`|Обновлены и дополнены:<br><ul><li><p>методы `getSelectedDataAsync` и `setSelectedDataAsync` для поддержки считывания выделенного пользователем фрагмента и его замены в теме и тексте сообщения или встречи;</p></li><li><p>методы `displayReplyAllForm` и `displayReplyForm` для поддержки добавления вложения в форму ответа для встречи.</p></li></ul>|Mailbox 1.2|
|`Item`|Обновлен для включения методов и полей для создания надстроек Outlook, активирующихся в режиме создания. |1.1|
|`Binding`|Обновлен для поддержки привязки к таблице в контентных надстройках для Access.|1.1|
|`Bindings`|Обновлен для поддержки привязки к таблице в контентных надстройках для Access.|1.1|
|`Body`|Добавлен для поддержки создания и изменения текста сообщения или встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|
|`Document`|Обновлены и дополнены: <ul><li><p>Поддержка свойств <a href="/javascript/api/office/office.document" target="_blank">mode</a>, <a href="/javascript/api/office/office.document#settings" target="_blank">settings</a>, а также <a href="/javascript/api/office/office.document" target="_blank">url</a> в контентных надстройках для Access.</p></li><li><p>Получение документа в виде PDF-файла с помощью метода <a href="/javascript/api/office/office.document#getfileasync-filetype--options--callback-" target="_blank">getFileAsync</a> в надстройках для PowerPoint и Word.</p></li><li><p>Получение свойств файла с помощью метода <a href="/javascript/api/office/office.document#getfilepropertiesasync-options--callback-" target="_blank">getFileProperties</a> в надстройках для Excel, PowerPoint и Word.</p></li><li><p>Переход к расположениям и объектам в документе с помощью метода <a href="/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-" target="_blank">goToByldAsync</a> в надстройках для Excel и PowerPoint.</p></li><li><p>Получение идентификатора, заголовка и индекса выбранных слайдов с помощью метода <a href="/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-" target="_blank">getSelectedDataAsync</a> (при указании нового перечисления <span class="keyword">Office.CoercionType.SlideRange</span><a href="/javascript/api/office/office.coerciontype" target="_blank">coercionType</a>) в надстройках для PowerPoint.</p></li></ul>|1.1|
|`Location`|Добавлен, чтобы стало возможным задание место встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|
|`Office`|Обновлен метод select для поддержки получения привязки в контентных надстройках для Access.|1.1|
|`Recipients`|Добавлен для поддержки получения и установки получателей сообщения или встречи в приложениях режима создания.|1.1|
|`Settings`|Обновлен для поддержки создания пользовательских настроек в контентных надстройках для Access.|1.1|
|`Subject`|Добавлен, чтобы стало возможным получение и задание темы сообщения или встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|
|`Time`|Добавлен, чтобы стало возможным получение и задание времени начала и окончания встречи в надстройках Outlook, активирующихся в режиме создания.|1.1|

### <a name="new-and-updated-enumerations"></a>Новые и обновленные перечисления

|**Объект**|**Описание**|**Версия**|
|:-----|:-----|:-----|
|`ActiveView`|Указывает состояние активного представления документа, например возможность редактирования документа пользователем. Добавлен, чтобы надстройки для PowerPoint могли определить, просматривают ли пользователи презентацию (**Показ слайдов**) или редактируют слайды. |1.1|
|`CoercionType`|Добавлен элемент **Office.CoercionType.SlideRange** для поддержки получения выбранного диапазона слайдов с помощью метода **getSelectedDataAsync** в надстройках для PowerPoint.|1.1|
|`EventType`|Добавлено новое событие ActiveViewChanged.|1.1|
|`FileType`|Добавлена возможность указания выходного файла в формате PDF.|1.1|
|`GoToType`|Добавлен для указания места или объекта в документе, к которому необходимо перейти.|1.1|

