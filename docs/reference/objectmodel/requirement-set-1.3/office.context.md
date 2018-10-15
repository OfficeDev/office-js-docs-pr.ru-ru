
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

### <a name="namespaces"></a>Пространства имен

[mailbox](office.context.mailbox.md): предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.

### <a name="members"></a>Члены

####  <a name="displaylanguage-string"></a>displayLanguage :String

Получает определенный пользователем языковой стандарт в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.

Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.

##### <a name="type"></a>Тип:

*   Строка

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="officetheme-object"></a>officeTheme :Object

Предоставляет доступ к свойствам цветов темы Office.

> [!NOTE]
> Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.

Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office**, и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.

##### <a name="type"></a>Тип:

*   Object

##### <a name="properties"></a>Свойства:

|Имя| Тип| Описание|
|---|---|---|
|`bodyBackgroundColor`| String|Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.|
|`bodyForegroundColor`| String|Получает основной цвет текста сообщения для темы Office в виде шестнадцатеричной триады цветов.|
|`controlBackgroundColor`| String|Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.|
|`controlForegroundColor`| String|Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.|

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|

##### <a name="example"></a>Пример

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)

Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.

Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.

##### <a name="type"></a>Тип:

*   [RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a>Требования

|Требование| Значение|
|---|---|
|[Версия минимального набора требований для почтового ящика (mailbox)](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Минимальный уровень разрешений](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Ограниченный доступ|
|[Применимый режим Outlook](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Создание или чтение|