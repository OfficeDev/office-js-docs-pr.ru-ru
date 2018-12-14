---
title: Обход правила ограничения домена в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: e5aa329eb3f073f3544d8446683debed3239fd00
ms.sourcegitcommit: 3d8454055ba4d7aae12f335def97357dea5beb30
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/14/2018
ms.locfileid: "27270602"
---
# <a name="addressing-same-origin-policy-limitations-in-office-add-ins"></a><span data-ttu-id="b286e-102">Обход правила ограничения домена в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="b286e-102">Addressing same-origin policy limitations in Office Add-ins</span></span>


<span data-ttu-id="b286e-p101">Правило ограничения домена, применяемое браузером, не позволяет скрипту, загруженному из одного домена, получать и обрабатывать свойства веб-страницы из другого домена. Это значит, что по умолчанию домен запрошенного URL-адреса должен быть тем же, что и домен текущей веб-страницы. Например, это правило не позволяет веб-странице в одном домене выполнять вызовы веб-службы [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) в другом домене.</span><span class="sxs-lookup"><span data-stu-id="b286e-p101">The same-origin policy enforced by the browser prevents a script loaded from one domain from getting or manipulating properties of a webpage from another domain. This means that, by default, the domain of a requested URL must be the same as the domain of the current webpage. For example, this policy will prevent a webpage in one domain from making [XmlHttpRequest](https://www.w3.org/TR/XMLHttpRequest/) web-service calls to a domain other than the one where it is hosted.</span></span>

<span data-ttu-id="b286e-106">Так как надстройки Office размещаются в браузере, правило ограничения домена также применяется к скриптам, работающим на веб-страницах этих надстроек.</span><span class="sxs-lookup"><span data-stu-id="b286e-106">Because Office Add-ins are hosted in a browser control, the same-origin policy applies to script running in their web pages as well.</span></span>

<span data-ttu-id="b286e-107">Чтобы обойти реализацию правила ограничения домена при разработке надстроек, можно использовать следующие способы.</span><span class="sxs-lookup"><span data-stu-id="b286e-107">To overcome same-origin policy enforcement when you develop add-ins, you can:</span></span>

- <span data-ttu-id="b286e-108">Использование JSON/P для анонимного доступа.</span><span class="sxs-lookup"><span data-stu-id="b286e-108">Use JSON/P for anonymous access.</span></span> 
    
- <span data-ttu-id="b286e-109">Реализация скриптов на стороне сервера с использованием схемы проверки подлинности на основе маркеров.</span><span class="sxs-lookup"><span data-stu-id="b286e-109">Implement server-side script using a token-based authentication scheme.</span></span>
    
- <span data-ttu-id="b286e-110">Использование CORS.</span><span class="sxs-lookup"><span data-stu-id="b286e-110">Using cross-origin resource sharing (CORS).</span></span>
    
- <span data-ttu-id="b286e-111">Создание собственного прокси-сервера с использованием IFRAME и POST MESSAGE.</span><span class="sxs-lookup"><span data-stu-id="b286e-111">Build your own proxy using IFRAME and POST MESSAGE.</span></span>
    

## <a name="using-jsonp-for-anonymous-access"></a><span data-ttu-id="b286e-112">Использование JSON/P для анонимного доступа</span><span class="sxs-lookup"><span data-stu-id="b286e-112">Using JSON/P for anonymous access</span></span>


<span data-ttu-id="b286e-p102">Это ограничение можно обойти, используя JSON/P, чтобы указать прокси-сервер для веб-службы. Для этого можно включить тег `script` с атрибутом `src`, указывающим на некоторый сценарий, расположенный на любом домене. Вы можете создать теги `script` программным способом, динамически создать URL-адрес, на который будет указывать атрибут `src`, а затем передать параметры по URL-адресу с помощью параметров запроса URI. Поставщики веб-служб создают и размещают код JavaScript с использованием определенных URL-адресов и возвращают разные сценарии в зависимости от параметров запроса URI. Затем эти сценарии выполняются в точке вставки и работают надлежащим образом.</span><span class="sxs-lookup"><span data-stu-id="b286e-p102">One way to overcome this limitation is to use JSON/P to provide a proxy for the web service. You do this by including a `script` tag with a `src` attribute that points to some script hosted on any domain. You can programmatically create the `script` tags, dynamically create the URL to point the `src` attribute to, and then pass parameters to the URL via URI query parameters. Web service providers create and host JavaScript code at specific URLs, and return different scripts depending on the URI query parameters. These scripts then execute where they are inserted and work as expected.</span></span>

<span data-ttu-id="b286e-118">Ниже приведен пример кода JSON/P, где используется метод, который будет работать в любых надстройках Office.</span><span class="sxs-lookup"><span data-stu-id="b286e-118">The following is an example of JSON/P that uses a technique that will work in any Office Add-in.</span></span>

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


## <a name="implementing-server-side-script-using-a-token-based-authentication-scheme"></a><span data-ttu-id="b286e-119">Реализация сценариев на стороне сервера с использованием схемы проверки подлинности на основе маркеров</span><span class="sxs-lookup"><span data-stu-id="b286e-119">Implementing server-side script using a token-based authentication scheme</span></span>


<span data-ttu-id="b286e-120">Другой способ устранения ограничений, связанных с правилом ограничения домена, состоит в реализации веб-страницы надстройки как страницы ASP, использующей OAuth или выполняющей кэширование учетных данных в файлах cookie.</span><span class="sxs-lookup"><span data-stu-id="b286e-120">Another way to address same-origin policy limitations is to implement the add-in's webpage as an ASP page that uses OAuth or caches credentials in cookies.</span></span>

<span data-ttu-id="b286e-121">Пример кода на стороне сервера, демонстрирующего использование объекта `Cookie` в `System.Net` для получения и задания значений файлов cookie, см. в свойстве [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2).</span><span class="sxs-lookup"><span data-stu-id="b286e-121">For an example of server-side code that shows how to use the  `Cookie` object in `System.Net` to get and set cookie values, see the [Value](https://docs.microsoft.com/dotnet/api/system.net.cookie.value?view=netframework-4.7.2) property.</span></span>


## <a name="using-cross-origin-resource-sharing-cors"></a><span data-ttu-id="b286e-122">Использование CORS</span><span class="sxs-lookup"><span data-stu-id="b286e-122">Using cross-origin resource sharing (CORS)</span></span>


<span data-ttu-id="b286e-123">Пример использования функций предоставления общего доступа к ресурсам разного происхождения [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html) см. в разделе Cross Origin Resource Sharing (CORS) статьи [Новые возможности XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span><span class="sxs-lookup"><span data-stu-id="b286e-123">For an example of using the cross-origin resource sharing feature of [XmlHttpRequest2](https://dvcs.w3.org/hg/xhr/raw-file/tip/Overview.html), see the "Cross Origin Resource Sharing (CORS)" section of [New Tricks in XMLHttpRequest2](https://www.html5rocks.com/en/tutorials/file/xhr2/).</span></span>


## <a name="building-your-own-proxy-using-iframe-and-post-message"></a><span data-ttu-id="b286e-124">Создание собственного прокси-сервера с использованием IFRAME и POST MESSAGE</span><span class="sxs-lookup"><span data-stu-id="b286e-124">Building your own proxy using IFRAME and POST MESSAGE</span></span>


<span data-ttu-id="b286e-125">Пример создания собственного прокси-сервера с использованием IFRAME и POST MESSAGE см. в статье [Обмен сообщениями между окнами](http://ejohn.org/blog/cross-window-messaging/).</span><span class="sxs-lookup"><span data-stu-id="b286e-125">For an example of how to build your own proxy using IFRAME and POST MESSAGE, see [Cross-Window Messaging](http://ejohn.org/blog/cross-window-messaging/).</span></span>


## <a name="see-also"></a><span data-ttu-id="b286e-126">См. также</span><span class="sxs-lookup"><span data-stu-id="b286e-126">See also</span></span>

- [<span data-ttu-id="b286e-127">Конфиденциальность и безопасность надстроек Office</span><span class="sxs-lookup"><span data-stu-id="b286e-127">Privacy and security for Office Add-ins</span></span>](../concepts/privacy-and-security.md)
    
