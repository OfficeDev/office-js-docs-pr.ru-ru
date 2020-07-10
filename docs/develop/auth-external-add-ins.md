---
title: Авторизация внешних служб в надстройке Office
description: Получение авторизации для сторонних источников данных (отличных от Майкрософт), например Google, Facebook, LinkedIn, SalesForce и GitHub, с помощью OAuth 2.0, кода авторизации и неявных потоков.
ms.date: 08/07/2019
localization_priority: Normal
ms.openlocfilehash: fd180e11106e7e1e2f20f539746535c4310ad81e
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093744"
---
# <a name="authorize-external-services-in-your-office-add-in"></a>Авторизация внешних служб в надстройке Office

Popular online services, including Microsoft 365, Google, Facebook, LinkedIn, SalesForce, and GitHub, let developers give users access to their accounts in other applications. This gives you the ability to include these services in your Office Add-in.

> [!NOTE]
> В оставшейся части этой статьи рассматривается доступ к сторонним службам (отличным от Майкрософт). Сведения о доступе к Microsoft Graph (в том числе Microsoft 365) можно найти в [статье доступ к Microsoft Graph с помощью единого входа](overview-authn-authz.md#access-to-microsoft-graph-with-sso) и [доступа к Microsoft Graph без единого входа](overview-authn-authz.md#access-to-microsoft-graph-without-sso).

The industry standard framework for enabling web application access to an online service is **OAuth 2.0**. In most situations, you don't need to know the details of how the framework works to use it in your add-in. Many libraries are available that simplify the details for you.

Основной принцип работы OAuth заключается в том, что приложение может быть [субъектом безопасности](/windows/security/identity-protection/access-control/security-principals) для самого себя (как пользователь или группа) и использовать собственное удостоверение и набор разрешений. В большинстве типичных сценариев, когда пользователь выполняет действие в надстройке Office, требующее вовлечения веб-службы, надстройка отправляет службе запрос на получение определенного набора разрешений для учетной записи пользователя. Затем служба предлагает пользователю предоставить надстройке эти разрешения. После предоставления разрешений служба отправляет надстройке небольшой зашифрованный *маркер доступа*. Надстройка может использовать службу, включая этот маркер во все свои запросы к API-интерфейсам службы. Но надстройка может действовать только в рамках разрешений, предоставленных ей пользователем. Кроме того, срок действия маркера ограничен указанным периодом времени.

Разные шаблоны OAuth (*потоки* или *типы предоставления*) предназначены для разных сценариев. Наиболее часто используются следующие два шаблона:

- **Неявный поток.** Обмен данными между надстройкой и веб-службой реализуется с помощью JavaScript на стороне клиента. Этот поток часто используется в одностраничных приложениях (SPA).
- **Поток кода авторизации**. Используется подключение типа *сервер-сервер* между веб-приложением надстройки и веб-службой. Следовательно, используется серверный код.

Поток OAuth предназначен для безопасной идентификации и авторизации приложения. В потоке кода авторизации вам предоставляется *секрет клиента*, который необходимо держать в тайне. Приложение без серверной части, например SPA, не может защитить секрет, поэтому в таких приложениях рекомендуется использовать неявный поток.

Дополнительные сведения о преимуществах и недостатках неявного потока и потока кода авторизации см. в разделах [Код авторизации](https://tools.ietf.org/html/rfc6749#section-1.3.1) и [Неявный поток](https://tools.ietf.org/html/rfc6749#section-1.3.2).

> [!NOTE]
> You also have the option of using a middleman service to perform authorization and pass the access token to your add-in. For details about this scenario, see the **Middleman services** section later in this article.

## <a name="using-the-implicit-flow-in-office-add-ins"></a>Использование неявного потока в надстройках Office

Узнать, поддерживает ли веб-служба неявный поток, можно из документации к ней.

Сведения о библиотеках, поддерживающих неявный поток, см. в разделе **Библиотеки** далее в этой статье.

## <a name="using-the-authorization-code-flow-in-office-add-ins"></a>Использование потока кода авторизации в надстройках Office

Many libraries are available for implementing the Authorization Code flow in various languages and frameworks. For more information about some of these libraries, see the **Libraries** section later in this article.

## <a name="libraries"></a>Библиотеки

Libraries are available for many languages and platforms, for both the Implicit flow and the Authorization Code flow. Some libraries are general purpose, while others are for specific online services.

**Google**: Search [GitHub.com/Google](https://github.com/google) for "auth" or the name of your language. Most of the relevant repos are named `google-auth-library-[name of language]`.

Если вы используете **Facebook**, выполните на сайте [Facebook для разработчиков](https://developers.facebook.com) поиск по запросу "library" или "sdk".

**General OAuth 2.0**: A page of links to libraries for over a dozen languages is maintained by the IETF OAuth Working Group at: [OAuth Code](https://oauth.net/code/). Note that some of these libraries are for implementing an OAuth compliant service. The libraries of interest to you as a an add-in developer are called *client* libraries on this page because your web server is a client of the OAuth compliant service.

## <a name="middleman-services"></a>Службы-посредники

Your add-in can use a middleman service such as [OAuth.io](https://oauth.io) or [Auth0](https://auth0.com) to perform authorization. A middleman service may either provide access tokens for popular online services or simplify the process of enabling social login for your add-in, or both. With very little code, your add-in can use either client-side script or server-side code to connect to the middleman service and it will send your add-in any required tokens for the online service. All of the authorization implementation code is in the middleman service. 

Рекомендуем сделать так, чтобы надстройка при проверке подлинности или авторизации использовала наши Dialog API для открытия страницы входа. Дополнительные сведения см. в статье [Использование Dialog API в потоке проверки подлинности](dialog-api-in-office-add-ins.md#use-the-dialog-apis-in-an-authentication-flow). Когда вы открываете диалоговое окно Office так, оно открывается в совершенно новом экземпляре браузера. При этом используется модуль JavaScript из экземпляра на родительской странице (например, область задач надстройки, FunctionFile). Токен и другие данные, которые можно преобразовать в строку, передается на родительскую страницу с помощью API под названием `messageParent`. Затем родительская страница может использовать этот токен для авторизованных вызовов ресурса. Из-за особенностей архитектуры те API, которые предоставляет служба-посредник, следует использовать с осторожностью. Часто служба предоставляет набор API, в котором ваш код создает определенный объект контекста, получающий токен и использующий его для последующих вызовов ресурса. Часто у службы есть один метод API, который делает начальный вызов *и* создает объект контекста. Подобный объект невозможно полностью преобразовать в строку, поэтому его нельзя передать из диалогового окна Office на родительскую страницу. Как правило, служба-посредник предоставляет второй набор API с более низким уровнем абстракции (например, REST API). Этот второй набор содержит API, который получает токен из службы, и другие API, которые передают токен службе при использовании для авторизованного доступа к ресурсу. Вам нужно работать с API на этом нижнем уровне абстракции, чтобы можно было получить токен в диалоговом окне Office, а затем использовать `messageParent` для его передачи на родительскую страницу. 

## <a name="what-is-cors"></a>Что такое CORS?

CORS stands for [Cross Origin Resource Sharing](https://developer.mozilla.org/docs/Web/HTTP/Access_control_CORS). For information about how to use CORS inside add-ins, see [Addressing same-origin policy limitations in Office Add-ins](addressing-same-origin-policy-limitations.md).
