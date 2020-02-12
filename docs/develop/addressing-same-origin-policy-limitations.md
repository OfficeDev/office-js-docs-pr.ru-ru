---
title: Обход правила ограничения домена в надстройках Office
description: ''
ms.date: 10/17/2019
localization_priority: Normal
ms.openlocfilehash: 2a47339bd5cc0b0bf919152b7078d5373382124f
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950448"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="c8d19-102">Обход правила ограничения домена в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="c8d19-102">Addressing same-origin policy limitations in Office Add-ins</span></span>

<span data-ttu-id="c8d19-p101">Правило ограничения домена, применяемое браузером, не позволяет скрипту, загруженному из одного домена, получать и обрабатывать свойства веб-страницы из другого домена. Это значит, что по умолчанию домен запрошенного URL-адреса должен быть тем же, что и домен текущей веб-страницы. Например, это правило не позволяет веб-странице в одном домене выполнять вызовы веб-службы [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) в другом домене.</span><span class="sxs-lookup"><span data-stu-id="c8d19-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="c8d19-106">Так как надстройки Office размещаются в браузере, правило ограничения домена также применяется к скриптам, работающим на веб-страницах этих надстроек.</span><span class="sxs-lookup"><span data-stu-id="c8d19-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="c8d19-107">Правило ограничения домена может стать помехой во многих случаях (например, если веб-приложение размещает контент и API на нескольких поддоменах).</span><span class="sxs-lookup"><span data-stu-id="c8d19-107">The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains.</span></span> <span data-ttu-id="c8d19-108">Существует несколько распространенных способов безопасного обхода правила ограничения домена.</span><span class="sxs-lookup"><span data-stu-id="c8d19-108">There are a few common techniques for securely overcoming same-origin policy enforcement.</span></span> <span data-ttu-id="c8d19-109">В этой статье предоставлены только короткие общие сведения о некоторых из них.</span><span class="sxs-lookup"><span data-stu-id="c8d19-109">This article can only provide the briefest introduction to some of them.</span></span> <span data-ttu-id="c8d19-110">Воспользуйтесь приведенными ссылками, чтобы приступить к изучению этих приемов.</span><span class="sxs-lookup"><span data-stu-id="c8d19-110">Please use the links provided to get started in your research of these techniques.</span></span>

## <a name="use-jsonp-for-anonymous-access"></a><span data-ttu-id="c8d19-111">Использование JSONP для анонимного доступа</span><span class="sxs-lookup"><span data-stu-id="c8d19-111">Use JSONP for anonymous access</span></span>

<span data-ttu-id="c8d19-112">Один из способов обойти правило ограничения домена — использовать [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp), чтобы указать прокси для веб-службы.</span><span class="sxs-lookup"><span data-stu-id="c8d19-112">One way to overcome same-origin policy limitations is to use [JSONP](https://www.w3schools.com/js/js_json_jsonp.asp) to provide a proxy for the web service.</span></span> <span data-ttu-id="c8d19-113">Для этого можно включить тег `script` с атрибутом `src`, указывающим на некоторый скрипт, размещенный на каком-либо домене.</span><span class="sxs-lookup"><span data-stu-id="c8d19-113">You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain.</span></span> <span data-ttu-id="c8d19-114">Вы можете создать теги `script` программным способом, динамически создать URL-адрес, на который будет указывать атрибут `src`, а затем передать параметры по URL-адресу с помощью параметров запроса URI.</span><span class="sxs-lookup"><span data-stu-id="c8d19-114">You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="c8d19-115">Поставщики веб-служб создают и размещают код JavaScript с использованием определенных URL-адресов и возвращают разные сценарии в зависимости от параметров запроса URI.</span><span class="sxs-lookup"><span data-stu-id="c8d19-115">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="c8d19-116">Затем эти сценарии выполняются в точке вставки и работают надлежащим образом.</span><span class="sxs-lookup"><span data-stu-id="c8d19-116">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="c8d19-117">Ниже приведен пример кода JSONP, где используется способ, который будет работать в любых надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="c8d19-117">The following is an example of JSONP that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a><span data-ttu-id="c8d19-118">Реализация серверного кода с использованием схемы авторизации на основе маркеров</span><span class="sxs-lookup"><span data-stu-id="c8d19-118">Implement server-side code using a token-based authorization scheme</span></span>

<span data-ttu-id="c8d19-119">Еще один способ обойти правило ограничения домена — предоставить серверный код, использующий потоки [OAuth 2.0](https://oauth.net/2/), чтобы обеспечить для одного домена авторизованный доступ к ресурсам, размещенным на другом.</span><span class="sxs-lookup"><span data-stu-id="c8d19-119">Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.</span></span> 


## <a name="use-cross-origin-resource-sharing-cors"></a><span data-ttu-id="c8d19-120">Совместное использование ресурсов из разных источников (CORS)</span><span class="sxs-lookup"><span data-stu-id="c8d19-120">Use cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="c8d19-121">Пример использования функций предоставления общего доступа к ресурсам разного происхождения [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) см. в разделе Cross Origin Resource Sharing (CORS) статьи [Новые возможности XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="c8d19-121">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a><span data-ttu-id="c8d19-122">Создание собственного прокси с использованием IFRAME и POST MESSAGE (обмен сообщениями между окнами)</span><span class="sxs-lookup"><span data-stu-id="c8d19-122">Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)</span></span>


<span data-ttu-id="c8d19-123">Пример создания собственного прокси с использованием IFRAME и POST MESSAGE см. в статье [Обмен сообщениями между окнами](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="c8d19-123">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="c8d19-124">См. также</span><span class="sxs-lookup"><span data-stu-id="c8d19-124">See also</span></span>

- [<span data-ttu-id="c8d19-125">Конфиденциальность и безопасность надстроек Office</span><span class="sxs-lookup"><span data-stu-id="c8d19-125">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
