# <a name="functionfile-element"></a>Элемент FunctionFile

Указывает файл с исходным кодом для операций, доступных через те команды надстройки, для выполнения которых используется функция JavaScript, а не отображается пользовательский интерфейс. Элемент **FunctionFile** является дочерним для [DesktopFormFactor](desktopformfactor.md) или [MobileFormFactor](mobileformfactor.md). Атрибуту **resid** элемента **FunctionFile** присваивается значение атрибута **id** элемента **Url** в элементе **Resources**. Последний содержит URL-адрес HTML-файла, который содержит или загружает все функции JavaScript, используемые для выполнения команд надстройки без пользовательского интерфейса, как определено элементом [Control](control.md).

Ниже приведен пример элемента **FunctionFile**.

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

JavaScript в HTML-файле, указанном элементом **FunctionFile** должен вызвать метод `Office.initialize` и определить именованные функции, которые принимают единственный параметр: `event`. Функции должны использовать API `item.notificationMessages`, чтобы предоставлять пользователю индикацию хода выполнения и сообщать об успехах или сбоях. Функция должна также вызвать `event.completed` при завершении выполнения. Имя функции используется в элементе **FunctionName** для кнопок без пользовательского интерфейса.

Ниже приведен пример HTML-файла для определения функции **trackMessage**.

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

Приведенный ниже пример кода показывает, как реализовать функцию, используемую элементом **FunctionName**.

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> Вызов метода **event.completed** означает, что событие успешно обработано. Если функция вызывается несколько раз, например при многократном выборе одной команды надстройки, все события автоматически помещаются в очередь. Первое событие запускается автоматически, а другие ожидают в очереди. Когда функция вызывает метод **event.completed**, для нее запускается следующий вызов в очереди. Вы должны вызывать метод **event.completed**. В противном случае функция не будет запускаться.