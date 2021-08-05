---
title: Разработка надстройки Office для работы с ITP при использовании сторонних файлов cookie
description: Работа с ITP и Office надстройки при использовании сторонних файлов cookie
ms.date: 07/8/2021
localization_priority: Normal
ms.openlocfilehash: 6a9452f24cb1cbd76c4f6cc3f39fab1f9310ec97
ms.sourcegitcommit: e570fa8925204c6ca7c8aea59fbf07f73ef1a803
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/05/2021
ms.locfileid: "53773477"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Разработка надстройки Office для работы с ITP при использовании сторонних файлов cookie

Если для Office надстройки требуются сторонние файлы cookie, эти файлы cookie будут заблокированы, если интеллектуальная профилактика отслеживания (ITP) используется временем запуска браузера, загрузив надстройку. Для проверки подлинности пользователей или для других сценариев, таких как хранение параметров, можно использовать сторонние файлы cookie.

Если ваша Office надстройка и веб-сайт должны полагаться на сторонние файлы cookie, используйте следующие действия для работы с ITP.

1. Настройка [авторизации OAuth 2.0](https://tools.ietf.org/html/rfc6749)таким образом, чтобы домен проверки подлинности (в вашем случае стороннее стороннее, ожидающее файлов cookie) перенародил маркер авторизации на   ваш веб-сайт. Используйте маркер для создания сеанса входа с помощью сервера Secure и [cookie HttpOnly.](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies)
2. Используйте [API служба хранилища доступа,](https://webkit.org/blog/8124/introducing-storage-access-api/)чтобы сторонние стороны могли запрашивать разрешения на доступ к его первому   участнику cookie. Текущие версии Office Mac и Office в Интернете поддерживают этот API.
    > [!NOTE]
    > Если вы используете файлы cookie для других целей, кроме проверки подлинности, то рассмотрите возможность `localStorage` использования для вашего сценария.

В следующем примере кода показано, как использовать API служба хранилища доступа.

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

## <a name="about-itp-and-third-party-cookies"></a>О ИТП и сторонних файлах cookie

Сторонние файлы cookie — это файлы cookie, загружаются в iframe, где домен отличается от кадра верхнего уровня. ItP может повлиять на сложные сценарии проверки подлинности, когда диалоговое окно всплывающее окно используется для ввода учетных данных, а затем доступ к файлам cookie необходим надстройке iframe для завершения потока проверки подлинности. ItP также может повлиять на сценарии бесшумной проверки подлинности, где ранее для проверки подлинности использовался диалоговое окно всплывающее окно, но после этого использование надстройки пытается проверить подлинность через скрытый iframe.

При разработке Office надстроек на Mac доступ к сторонним файлам cookie блокируется SDK MacOS Big Sur. Это происходит из-за того, что ИТП WKWebView включен по умолчанию в браузере Safari, а WKWebView блокирует все сторонние файлы cookie. Office mac версии 16.44 или более поздней версии интегрирован с MacOS Big Sur SDK.

В браузере Safari конечные пользователи могут переключать контрольный ящик **Prevent cross-site tracking** under **Preference**  >  **Privacy,** чтобы отключить ITP. Однако itP нельзя отключить для встроенного управления WKWebView.

## <a name="see-also"></a>См. также

- [Обработка ITP в Safari и других браузерах, где сторонние файлы cookie заблокированы](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Отслеживание предотвращения в WebKit](https://webkit.org/tracking-prevention/)
- ["Песочница конфиденциальности" Chrome](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Введение API служба хранилища доступа](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)