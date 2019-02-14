---
title: Обход правила ограничения домена в надстройках Office
description: ''
ms.date: 02/08/2019
localization_priority: Priority
ms.openlocfilehash: 52af2eef2881b48feb141182233bc194ae406aa0
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/13/2019
ms.locfileid: "29981994"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="766b5-102">Обход правила ограничения домена в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="766b5-102">Addressing same-origin policy limitations in Office Add-ins</span></span>

<span data-ttu-id="766b5-p101">Правило ограничения домена, применяемое браузером, не позволяет скрипту, загруженному из одного домена, получать и обрабатывать свойства веб-страницы из другого домена. Это значит, что по умолчанию домен запрошенного URL-адреса должен быть тем же, что и домен текущей веб-страницы. Например, это правило не позволяет веб-странице в одном домене выполнять вызовы веб-службы [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) в другом домене.</span><span class="sxs-lookup"><span data-stu-id="766b5-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="766b5-106">Так как надстройки Office размещаются в браузере, правило ограничения домена также применяется к скриптам, работающим на веб-страницах этих надстроек.</span><span class="sxs-lookup"><span data-stu-id="766b5-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="766b5-107">Правило ограничения домена может стать помехой во многих случаях (например, если веб-приложение размещает контент и API на нескольких поддоменах).</span><span class="sxs-lookup"><span data-stu-id="766b5-107">The same-origin policy can be an unnecessary handicap in many situations, such as when a web application hosts content and APIs across multiple subdomains.</span></span> <span data-ttu-id="766b5-108">Существует несколько распространенных способов безопасного обхода правила ограничения домена.</span><span class="sxs-lookup"><span data-stu-id="766b5-108">There are a few common techniques for securely overcoming same-origin policy enforcement.</span></span> <span data-ttu-id="766b5-109">В этой статье предоставлены только короткие общие сведения о некоторых из них.</span><span class="sxs-lookup"><span data-stu-id="766b5-109">This article can only provide the briefest introduction to some of them.</span></span> <span data-ttu-id="766b5-110">Воспользуйтесь приведенными ссылками, чтобы приступить к изучению этих приемов.</span><span class="sxs-lookup"><span data-stu-id="766b5-110">Please use the links provided to get started in your research of these techniques.</span></span>

## <a name="use-jsonp-for-anonymous-access"></a><span data-ttu-id="766b5-111">Использование JSON/P для анонимного доступа</span><span class="sxs-lookup"><span data-stu-id="766b5-111">Use JSON/P for anonymous access.</span></span>

<span data-ttu-id="766b5-112">Один из способов обойти правило ограничения домена — использовать [JSON/P](https://www.w3schools.com/js/js_json_jsonp.asp), чтобы указать прокси для веб-службы.</span><span class="sxs-lookup"><span data-stu-id="766b5-112">One way to overcome this limitation is to use JSON/P to provide a proxy for the web service.</span></span> <span data-ttu-id="766b5-113">Для этого можно включить тег `script` с атрибутом `src`, указывающим на некоторый скрипт, размещенный на каком-либо домене.</span><span class="sxs-lookup"><span data-stu-id="766b5-113">You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain.</span></span> <span data-ttu-id="766b5-114">Вы можете создать теги `script` программным способом, динамически создать URL-адрес, на который будет указывать атрибут `src`, а затем передать параметры по URL-адресу с помощью параметров запроса URI.</span><span class="sxs-lookup"><span data-stu-id="766b5-114">You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters.</span></span> <span data-ttu-id="766b5-115">Поставщики веб-служб создают и размещают код JavaScript с использованием определенных URL-адресов и возвращают разные сценарии в зависимости от параметров запроса URI.</span><span class="sxs-lookup"><span data-stu-id="766b5-115">Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters.</span></span> <span data-ttu-id="766b5-116">Затем эти сценарии выполняются в точке вставки и работают надлежащим образом.</span><span class="sxs-lookup"><span data-stu-id="766b5-116">These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="766b5-117">Ниже приведен пример кода JSON/P, где используется способ, который будет работать в любых надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="766b5-117">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implement-server-side-code-using-a-token-based-authorization-scheme"></a><span data-ttu-id="766b5-118">Реализация серверного кода с использованием схемы авторизации на основе маркеров</span><span class="sxs-lookup"><span data-stu-id="766b5-118">Implement server-side script using a token-based authentication scheme.</span></span>

<span data-ttu-id="766b5-119">Еще один способ обойти правило ограничения домена — предоставить серверный код, использующий потоки [OAuth 2.0](https://oauth.net/2/), чтобы обеспечить для одного домена авторизованный доступ к ресурсам, размещенным на другом.</span><span class="sxs-lookup"><span data-stu-id="766b5-119">Another way to address same-origin policy limitations is to provide server-side code that uses [OAuth 2.0](https://oauth.net/2/) flows to enable one domain to get authorized access to resources hosted on another.</span></span> 


## <a name="use-cross-origin-resource-sharing-cors"></a><span data-ttu-id="766b5-120">Совместное использование ресурсов из разных источников (CORS)</span><span class="sxs-lookup"><span data-stu-id="766b5-120">Using cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="766b5-121">Пример использования функций предоставления общего доступа к ресурсам разного происхождения [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) см. в разделе Cross Origin Resource Sharing (CORS) статьи [Новые возможности XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="766b5-121">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="build-your-own-proxy-using-iframe-and-post-message-cross-window-messaging"></a><span data-ttu-id="766b5-122">Создание собственного прокси с использованием IFRAME и POST MESSAGE (обмен сообщениями между окнами)</span><span class="sxs-lookup"><span data-stu-id="766b5-122">Build your own proxy using IFRAME and POST MESSAGE (Cross-Window Messaging)</span></span>


<span data-ttu-id="766b5-123">Пример создания собственного прокси с использованием IFRAME и POST MESSAGE см. в статье [Обмен сообщениями между окнами](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="766b5-123">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="766b5-124">См. также</span><span class="sxs-lookup"><span data-stu-id="766b5-124">See also</span></span>

- [<span data-ttu-id="766b5-125">Конфиденциальность и безопасность надстроек Office</span><span class="sxs-lookup"><span data-stu-id="766b5-125">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
