---
title: Разработка надстройки Office для работы с ITP при использовании сторонних файлов cookie
description: Работа с надстройкими ITP и Office при использовании сторонних файлов cookie
ms.date: 03/12/2021
localization_priority: Normal
ms.openlocfilehash: e66fc25e1dc0f3a93fdf38c1d0c099d3a68459d3
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178043"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a><span data-ttu-id="80cf6-103">Разработка надстройки Office для работы с ITP при использовании сторонних файлов cookie</span><span class="sxs-lookup"><span data-stu-id="80cf6-103">Develop your Office Add-in to work with ITP when using third-party cookies</span></span>

<span data-ttu-id="80cf6-104">Если надстройка Office требует сторонних файлов cookie, эти файлы cookie будут заблокированы, если интеллектуальная профилактика отслеживания (ITP) используется временем запуска браузера, загрузив надстройку.</span><span class="sxs-lookup"><span data-stu-id="80cf6-104">If your Office Add-in requires third-party cookies, those cookies are blocked if Intelligent Tracking Prevention (ITP) is used by the browser runtime that loaded your add-in.</span></span> <span data-ttu-id="80cf6-105">Для проверки подлинности пользователей или для других сценариев, таких как хранение параметров, можно использовать сторонние файлы cookie.</span><span class="sxs-lookup"><span data-stu-id="80cf6-105">You may be using third-party cookies to authenticate users, or for other scenarios, such as storing settings.</span></span>

<span data-ttu-id="80cf6-106">Если надстройка Office и веб-сайт должны полагаться на сторонние файлы cookie, используйте следующие действия для работы с ITP:</span><span class="sxs-lookup"><span data-stu-id="80cf6-106">If your Office Add-in and website must rely on third-party cookies, use the following steps to work with ITP:</span></span>

1. <span data-ttu-id="80cf6-107">Настройка [авторизации OAuth 2.0](https://tools.ietf.org/html/rfc6749)таким образом, чтобы домен проверки подлинности (в вашем случае стороннее стороннее, ожидающее файлов cookie) перенародил маркер авторизации на   ваш веб-сайт.</span><span class="sxs-lookup"><span data-stu-id="80cf6-107">Set up [OAuth 2.0 Authorization](https://tools.ietf.org/html/rfc6749) so that the authenticating domain (in your case, the third-party that expects cookies) forwards an authorization token to your website.</span></span> <span data-ttu-id="80cf6-108">Используйте маркер для создания сеанса входа с помощью сервера Secure и [cookie HttpOnly.](https://developer.mozilla.org/en-US/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)</span><span class="sxs-lookup"><span data-stu-id="80cf6-108">Use the token to establish a first-party login session with a server-set Secure and [HttpOnly cookie](https://developer.mozilla.org/en-US/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).</span></span>
2. <span data-ttu-id="80cf6-109">Используйте [API доступа](https://webkit.org/blog/8124/introducing-storage-access-api/)к хранилищам, чтобы сторонние стороны могли запрашивать разрешения на доступ к первому   участнику cookie.</span><span class="sxs-lookup"><span data-stu-id="80cf6-109">Use the [Storage Access API](https://webkit.org/blog/8124/introducing-storage-access-api/) so that the third-party can request permission to get access to its first-party cookies.</span></span> <span data-ttu-id="80cf6-110">Текущие версии Office на Mac и Office в Интернете поддерживают этот API.</span><span class="sxs-lookup"><span data-stu-id="80cf6-110">Current versions of Office on Mac and Office on the web both support this API.</span></span>
    > [!NOTE]
    > <span data-ttu-id="80cf6-111">Если вы используете файлы cookie для других целей, кроме проверки подлинности, то рассмотрите возможность `localStorage` использования для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="80cf6-111">If you're using cookies for purposes other than authentication, then consider using `localStorage` for your scenario.</span></span>

<span data-ttu-id="80cf6-112">В следующем примере кода показано, как использовать API доступа к хранилищам:</span><span class="sxs-lookup"><span data-stu-id="80cf6-112">The following code sample shows how to use the Storage Access API:</span></span>

```javascript
function displayLoginButton() {
  var button = createLoginButton();
  button.addEventListener("click", function(ev) {
    document.requestStorageAccess().then(function() {
      authenticateWithCookies(); 
    }).catch(function() {
      // User must have previously interacted with this domain loaded in a top frame
      // Also you should have previously written a cookie when domain was loaded in the top frame
      console.error("User cancelled or requirements were not met.");
    });
  });
}

if (document.hasStorageAccess) { 
  document.hasStorageAccess().then(function(hasStorageAccess) { 
    if (!hasStorageAccess) { 
      displayLoginButton(); 
    } else { 
      authenticateWithCookies(); 
    } 
  }); 
} else { 
    authenticateWithCookies(); 
} 
```

## <a name="about-itp-and-third-party-cookies"></a><span data-ttu-id="80cf6-113">О ИТП и сторонних файлах cookie</span><span class="sxs-lookup"><span data-stu-id="80cf6-113">About ITP and third-party cookies</span></span>

<span data-ttu-id="80cf6-114">Сторонние файлы cookie — это файлы cookie, загружаются в iframe, где домен отличается от кадра верхнего уровня.</span><span class="sxs-lookup"><span data-stu-id="80cf6-114">Third-party cookies are cookies that are loaded in an iframe, where the domain is different from the top level frame.</span></span> <span data-ttu-id="80cf6-115">ItP может повлиять на сложные сценарии проверки подлинности, когда диалоговое окно всплывающее окно используется для ввода учетных данных, а затем доступ к файлам cookie необходим надстройке iframe для завершения потока проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="80cf6-115">ITP could affect complex authentication scenarios, where a popup dialog is used to enter credentials and then the cookie access is needed by an add-in iframe to complete the authentication flow.</span></span> <span data-ttu-id="80cf6-116">ItP также может повлиять на сценарии бесшумной проверки подлинности, где ранее для проверки подлинности использовался диалоговое окно всплывающее окно, но после этого использование надстройки пытается проверить подлинность через скрытый iframe.</span><span class="sxs-lookup"><span data-stu-id="80cf6-116">ITP could also affect silent authentication scenarios, where you have previously used a popup dialog to authenticate, but subsequent use of the add-in tries to authenticate through a hidden iframe.</span></span>

<span data-ttu-id="80cf6-117">При разработке надстроек Office на Mac доступ к сторонним файлам cookie блокируется SDK MacOS Big Sur.</span><span class="sxs-lookup"><span data-stu-id="80cf6-117">When developing Office Add-ins on Mac, access to third-party cookies is blocked by the MacOS Big Sur SDK.</span></span> <span data-ttu-id="80cf6-118">Это происходит из-за того, что ITP WebKit включен по умолчанию в браузере Safari, а WKWebview блокирует все сторонние файлы cookie.</span><span class="sxs-lookup"><span data-stu-id="80cf6-118">This is because WebKit ITP is enabled by default on the Safari browser, and WKWebview blocks all third-party cookies.</span></span> <span data-ttu-id="80cf6-119">Office на Mac версии 16.44 или более поздней версии интегрирован с MacOS Big Sur SDK.</span><span class="sxs-lookup"><span data-stu-id="80cf6-119">Office on Mac version 16.44 or later is integrated with the MacOS Big Sur SDK.</span></span>

<span data-ttu-id="80cf6-120">В браузере Safari конечные пользователи могут переключать контрольный ящик **Prevent cross-site tracking** under **Preference**  >  **Privacy,** чтобы отключить ITP.</span><span class="sxs-lookup"><span data-stu-id="80cf6-120">In the Safari browser, end users can toggle the **Prevent cross-site tracking** checkbox under **Preference** > **Privacy** to turn off ITP.</span></span> <span data-ttu-id="80cf6-121">Однако itP нельзя отключить для встроенного управления WebKit2.</span><span class="sxs-lookup"><span data-stu-id="80cf6-121">However, ITP cannot be turned off for the embedded WebKit2 control.</span></span>

## <a name="see-also"></a><span data-ttu-id="80cf6-122">См. также</span><span class="sxs-lookup"><span data-stu-id="80cf6-122">See also</span></span>

- [<span data-ttu-id="80cf6-123">Обработка ITP в Safari и других браузерах, где сторонние файлы cookie заблокированы</span><span class="sxs-lookup"><span data-stu-id="80cf6-123">Handle ITP in Safari and other browsers where third-party cookies are blocked</span></span>](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [<span data-ttu-id="80cf6-124">Отслеживание предотвращения в WebKit</span><span class="sxs-lookup"><span data-stu-id="80cf6-124">Tracking Prevention in WebKit</span></span>](https://webkit.org/tracking-prevention/)
- [<span data-ttu-id="80cf6-125">"Песочница конфиденциальности" Chrome</span><span class="sxs-lookup"><span data-stu-id="80cf6-125">Chrome’s “Privacy Sandbox”</span></span>](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [<span data-ttu-id="80cf6-126">Введение API доступа к хранилищам</span><span class="sxs-lookup"><span data-stu-id="80cf6-126">Introducing the Storage Access API</span></span>](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)