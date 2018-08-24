---
title: Обход правила ограничения домена в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 536e02d2367bef81d4a6e49098d66833c99f5e50
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925110"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Обход правила ограничения домена в надстройках Office


Правило ограничения домена, применяемое браузером, не позволяет скрипту, загруженному из одного домена, получать и обрабатывать свойства веб-страницы из другого домена. Это значит, что по умолчанию домен запрошенного URL-адреса должен быть тем же, что и домен текущей веб-страницы. Например, это правило не позволяет веб-странице в одном домене выполнять вызовы веб-службы [XmlHttpRequest](http://www.w3.org/TR/XMLHttpRequest/) в другом домене.

Так как надстройки Office размещаются в браузере, правило ограничения домена также применяется к скриптам, работающим на веб-страницах этих надстроек.

Чтобы обойти реализацию правила ограничения домена при разработке надстроек, можно использовать следующие способы.

- Использование JSON/P для анонимного доступа. 
    
- Реализация скриптов на стороне сервера с использованием схемы проверки подлинности на основе маркеров.
    
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


Другой способ устранения ограничений, связанных с правилом ограничения домена, состоит в реализации веб-страницы надстройки как страницы ASP, использующей OAuth или выполняющей кэширование учетных данных в файлах cookie.

Пример кода на стороне сервера, демонстрирующего использование объекта `Cookie` в `System.Net` для получения и задания значений файлов cookie, см. в свойстве [Value](https://msdn.microsoft.com/library/4f772twc).


## <a name="using-cross-origin-resource-sharing-cors"></a>Использование CORS


Пример использования функций предоставления общего доступа к ресурсам разного происхождения [XmlHttpRequest2](http://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) см. в разделе Cross Origin Resource Sharing (CORS) статьи [Новые возможности XMLHttpRequest2](http://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a>Создание собственного прокси-сервера с использованием IFRAME и POST MESSAGE


Пример создания собственного прокси-сервера с использованием IFRAME и POST MESSAGE см. в статье [Обмен сообщениями между окнами](http://ejohn.org/blog/cross-window-messaging/).


## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек Office](../concepts/privacy-and-security.md)
    
