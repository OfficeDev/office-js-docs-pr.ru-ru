
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Решение ограничений политик одинакового происхождения в надстройках для Office


Политика единого происхождения, налагаемая браузером, препятствует сценарию, загруженному из одного домена, в получении и обработке свойств веб-страницы из другого домена. Это значит, что по умолчанию домен запрошенного URL-адреса должен быть тем же, что и домен текущей веб-страницы. Например, эта политика препятствует веб-странице в одном домене выполнять вызовы веб-службы [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) в домен, отличный от домена своего размещения.

Так как Надстройки Office размещаются в браузере, политика единого происхождения также применяется к сценариям, работающим на веб-страницах этих надстроек.

Чтобы обойти реализацию политики единого происхождения при разработке надстроек, можно использовать следующие способы.

- Использование JSON/P для анонимного доступа 
    
- Реализация сценариев на стороне сервера с использованием схемы проверки подлинности на основе маркеров
    
- Использование CORS.
    
- Создание собственного прокси-сервера с использованием IFRAME и POST MESSAGE.
    

## <a name="using-jsonp-for-anonymous-access"></a>Использование JSON/P для анонимного доступа


Это ограничение можно обойти, используя JSON/P, чтобы указать прокси-сервер для веб-службы. Для этого можно включить тег `script` с атрибутом `src`, указывающим на некоторый сценарий, расположенный на любом домене. Вы можете создать теги `script` программным способом, динамически создать URL-адрес, на который будет указывать атрибут `src`, а затем передать параметры по URL-адресу с помощью параметров запроса URI. Поставщики веб-служб создают и размещают код JavaScript с использованием определенных URL-адресов и возвращают разные сценарии в зависимости от параметров запроса URI. Затем эти сценарии выполняются в точке вставки и работают надлежащим образом.

Ниже приведен пример кода JSON/P, где используется метод, который будет работать в любых надстройках Office.

```js
// Dynamically create an HTML SCRIPT element that obtains the details for the specified video.
function loadVideoDetails(videoIndex) {
    // Dynamically create a new HTML SCRIPT element in the webpage.
    var script = document.createElement("script");
    // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
    // as the value of the src attribute of the SCRIPT element. 
    script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" + 
        videos[videoIndex].Id + "?alt=json-in-script&amp;callback=videoDetailsLoaded");
    // Insert the SCRIPT element at the end of the HEAD section.
    document.getElementsByTagName('head')[0].appendChild(script);
}

```


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a>Реализация сценариев на стороне сервера с использованием схемы проверки подлинности на основе маркеров


Другой способ устранения ограничений, связанных с политикой единого происхождения, состоит в реализации веб-страницы надстройки как страницы ASP, использующей OAuth или выполняющей кэширование учетных данных в файлах cookie.

Пример использования OAuth для проверки подлинности см. в статье [Веб-часть Twitter SharePoint с использованием OAuth](http://aidangarnish.net/post/Twitter-SharePoint-Web-Part-With-OAuth).

Пример кода на стороне сервера, демонстрирующего использование объекта `Cookie` в `System.Net` для получения и задания значений файлов cookie, см. в свойстве [Value](http://msdn2.microsoft.com/EN-US/library/4f772twc).


## <a name="using-cross-origin-resource-sharing-cors"></a>Использование CORS


Пример использования функций предоставления общего доступа к ресурсам разного происхождения [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) см. в разделе Cross Origin Resource Sharing (CORS) статьи [Новые возможности XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a>Создание собственного прокси-сервера с использованием IFRAME и POST MESSAGE


Пример создания собственного прокси-сервера с использованием IFRAME и POST MESSAGE см. в статье [Обмен сообщениями между окнами](http://ejohn.org/blog/cross-window-messaging/).


## <a name="additional-resources"></a>Дополнительные ресурсы


- [Конфиденциальность и безопасность надстроек для Office](../develop/privacy-and-security.md)
    
