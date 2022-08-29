---
title: Разработка надстройки Office для работы с ITP при использовании сторонних файлов cookie
description: Как работать с ITP и надстройки Office при использовании сторонних файлов cookie
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: b01051fa39441fddb2453b0bd95a0629ebf3ef65
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67423092"
---
# <a name="develop-your-office-add-in-to-work-with-itp-when-using-third-party-cookies"></a>Разработка надстройки Office для работы с ITP при использовании сторонних файлов cookie

Если надстройке Office требуются сторонние файлы cookie, эти файлы cookie блокируются, если среда выполнения[](../testing/runtimes.md), загрузив надстройку, использует интеллектуальную предотвращение отслеживания (ITP). Вы можете использовать сторонние файлы cookie для проверки подлинности пользователей или в других сценариях, таких как хранение параметров.

Если надстройка Office и веб-сайт должны полагаться на сторонние файлы cookie, выполните следующие действия для работы с ITP.

1. Настройка авторизации [OAuth 2.0](https://tools.ietf.org/html/rfc6749) таким образом, чтобы домен проверки подлинности (в вашем случае сторонний поставщик, который ожидает файлы cookie) перенаправлял маркер авторизации на ваш веб-сайт. Используйте маркер, чтобы установить сеанс входа в систему с помощью файлов cookie, заданных сервером Secure и [HttpOnly](https://developer.mozilla.org/docs/Web/HTTP/Cookies#Secure_and_HttpOnly_cookies).
1. Используйте [API доступа к хранилищу](https://webkit.org/blog/8124/introducing-storage-access-api/) , чтобы сторонние поставщики могли запрашивать разрешение на доступ к своим файлам cookie. Текущие версии Office для Mac и Office в Интернете поддерживают этот API.
    > [!NOTE]
    > Если вы используете файлы cookie для других целей, кроме проверки подлинности, рассмотрите возможность `localStorage` использования в вашем сценарии.

В следующем примере кода показано, как использовать API доступа к хранилищу.

```javascript
function displayLoginButton() {
  const button = createLoginButton();
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

## <a name="about-itp-and-third-party-cookies"></a>Сведения о ITP и сторонних файлах cookie

Сторонние файлы cookie — это файлы cookie, загружаемые в iframe, где домен отличается от фрейма верхнего уровня. ItP может повлиять на сложные сценарии проверки подлинности, в которых для ввода учетных данных используется всплывающее диалоговое окно, а затем для завершения потока проверки подлинности требуется доступ к файлу cookie в iframe надстройки. ItP также может повлиять на сценарии автоматической проверки подлинности, где вы ранее использовали всплывающее диалоговое окно для проверки подлинности, но последующее использование надстройки пытается выполнить проверку подлинности через скрытый iframe.

При разработке надстроек Office на Mac доступ к сторонним файлам cookie блокируется пакетом SDK для MacOS Big Sur. Это связано с тем, что WKWebView ITP включен по умолчанию в браузере Safari, а WKWebView блокирует все сторонние файлы cookie. Office для Mac версии 16.44 или более поздней интегрирован с пакетом SDK для MacOS Big Sur.

В браузере Safari конечные  >  пользователи могут установить флажок  "Запретить межсайтовое отслеживание" в разделе "Конфиденциальность предпочтений", чтобы отключить ITP. Однако itP нельзя отключить для внедренного элемента управления WKWebView.

## <a name="see-also"></a>См. также

- [Обработка ITP в Safari и других браузерах, где сторонние файлы cookie заблокированы](/azure/active-directory/develop/reference-third-party-cookies-spas)
- [Предотвращение отслеживания в WebKit](https://webkit.org/tracking-prevention/)
- [Песочница конфиденциальности Chrome](https://blog.chromium.org/2020/01/building-more-private-web-path-towards.html)
- [Знакомство с API доступа к хранилищу](https://blogs.windows.com/msedgedev/2020/07/08/introducing-storage-access-api/)
