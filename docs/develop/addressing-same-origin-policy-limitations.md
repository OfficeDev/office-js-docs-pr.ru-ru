---
title: Обход правила ограничения домена в надстройках Office
description: ''
ms.date: 02/08/2019
localization_priority: Priority
ms.openlocfilehash: 52af2eef2881b48feb141182233bc194ae406aa0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449073"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a>Обход правила ограничения домена в надстройках Office

Правило ограничения домена, применяемое браузером, не позволяет скрипту, загруженному из одного домена, получать и обрабатывать свойства веб-страницы из другого домена. Это значит, что по умолчанию домен запрошенного URL-адреса должен быть тем же, что и домен текущей веб-страницы. Например, это правило не позволяет веб-странице в одном домене выполнять вызовы веб-службы [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) в другом домене.

Так как надстройки Office размещаются в браузере, правило ограничения домена также применяется к скриптам, работающим на веб-страницах этих надстроек.

Правило ограничения домена может стать помехой во многих случаях (например, если веб-приложение размещает контент и API на нескольких поддоменах). Существует несколько распространенных способов безопасного обхода правила ограничения домена. В этой статье предоставлены только короткие общие сведения о некоторых из них. Воспользуйтесь приведенными ссылками, чтобы приступить к изучению этих приемов.

## <a name="use-jsonp-for-anonymous-access"></a>Использование JSON/P для анонимного доступа

Один из способов обойти правило ограничения домена — использовать [JSON/P](https://www.w3schools.com/js/js_json_jsonp.asp), чтобы указать прокси для веб-службы. Для этого можно включить тег `script` с атрибутом `src`, указывающим на некоторый скрипт, размещенный на каком-либо домене. Вы можете создать теги `script` программным способом, динамически создать URL-адрес, на который будет указывать атрибут `src`, а затем передать параметры по URL-адресу с помощью параметров запроса URI. Поставщики веб-служб создают и размещают код JavaScript с использованием определенных URL-адресов и возвращают разные сценарии в зависимости от параметров запроса URI. Затем эти сценарии выполняются в точке вставки и работают надлежащим образом.

Ниже приведен пример кода JSON/P, где используется способ, который будет работать в любых надстройках Office.

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a>Реализация серверного кода с использованием схемы авторизации на основе маркеров

Еще один способ обойти правило ограничения домена — предоставить серверный код, использующий потоки [OAuth 2.0](https://oauth.net/2/), чтобы обеспечить для одного домена авторизованный доступ к ресурсам, размещенным на другом. 


## <a name="use-cross-origin-resource-sharing-cors"></a>Совместное использование ресурсов из разных источников (CORS)


Пример использования функций предоставления общего доступа к ресурсам разного происхождения [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) см. в разделе Cross Origin Resource Sharing (CORS) статьи [Новые возможности XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a>Создание собственного прокси с использованием IFRAME и POST MESSAGE (обмен сообщениями между окнами)


Пример создания собственного прокси с использованием IFRAME и POST MESSAGE см. в статье [Обмен сообщениями между окнами](http://ejohn.org/blog/cross-window-messaging/).


## <a name="see-also"></a>См. также

- [Конфиденциальность и безопасность надстроек Office](../concepts/privacy-and-security.md)
    
