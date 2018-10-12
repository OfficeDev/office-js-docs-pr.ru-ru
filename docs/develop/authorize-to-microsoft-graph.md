---
title: Авторизованный доступ в Microsoft Graph из вашей надстройки Office
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 6d0b6f2002b71c4680b72d2f40492fff1abf15e2
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25505861"
---
# <a name="authorize-to-microsoft-graph-in-your-office-add-in-preview"></a>Авторизованный доступ в Microsoft Graph из вашей надстройки Office (предварительная версия)

Пользователи входят в Office (в Интернете, на мобильных устройствах и настольных компьютерах), используя личную учетную запись Майкрософт либо рабочую или учебную учетную запись (Office 365). Чтобы надстройка Office могла получить авторизованный доступ к [Microsoft Graph](https://developer.microsoft.com/graph/docs), лучше всего использовать учетные данные для входа пользователя в Office. Это позволяет пользователям получить доступ к своим данным Microsoft Graph без необходимости повторного входа. 

> [!NOTE]
> API единого входа в настоящее время поддерживается в предварительной версии для Word, Excel, Outlook и PowerPoint. Дополнительные сведения о поддержке API единого входа см. в статье [Наборы обязательных элементов IdentityAPI](https://docs.microsoft.com/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets?view=office-js). При работе с надстройкой Outlook необходимо включить современную проверку подлинности для клиента Office 365. Со сведениями о том, как это сделать, можно ознакомиться в статье [Exchange Online: как включить в клиенте современную проверку подлинности](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx).

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>Архитектура надстройки для единого входа и Microsoft Graph

Помимо страниц и кода JavaScript веб-приложения, в надстройке также должны размещаться (с тем же [полным доменном именем](https://docs.microsoft.com/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)) один или несколько веб-API, которые будут получать маркер доступа и отправлять запросы к Microsoft Graph.

Манифест надстройки содержит разметку, которая указывает, как надстройка регистрируется в конечной точке Azure Active Directory (Azure AD) версии 2.0, а также задает необходимые надстройке разрешения для Microsoft Graph.

### <a name="how-it-works-at-runtime"></a>Принцип работы во время выполнения

На следующей схеме изображен процесс входа в систему и получения доступа к Microsoft Graph.

![Схема единого входа](../images/sso-access-to-microsoft-graph.png)

1. Код JavaScript надстройки вызывает новый API Office.js — [](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference). Он указывает ведущему приложению Office, что необходимо получить маркер доступа к надстройке. (Здесь и далее он будет называться **маркером доступа к начальной загрузке**, поскольку в этом процессе он позже будет заменен вторым маркером. Пример декодированного маркера доступа к начальной загрузке см. в разделе [Пример маркера доступа](sso-in-office-add-ins.md#example-access-token).)
1. Если пользователь не выполнил вход в Office, в ведущем приложении открывается всплывающее окно, в котором ему предлагается войти.
1. Если текущий пользователь запускает надстройку в первый раз, ему предлагается дать согласие.
1. Ведущее приложение Office запрашивает **маркер доступа к начальной загрузке** у конечной точки Azure AD версии 2.0 для текущего пользователя.
1. Azure AD отправляет маркер начальной загрузки ведущему приложению Office.
1. Ведущее приложение Office отправляет **маркер доступа к начальной загрузке** надстройке в составе объекта результата, возвращенного при вызове метода `getAccessTokenAsync`.
1. Код JavaScript надстройки отправляет HTTP-запрос к веб-API с тем же полным доменным именем, что и у надстройки. Этот запрос включает **маркер доступа к начальной загузке** в качестве подтверждения авторизации.  
1. Серверный код проверяет входящий **маркер доступа к начальной загрузке**.
1. Серверный код использует поток "от имени»" (определенный в [Обмен токенами OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) и [демон или серверное приложение для сценария веб-API Azure](https://docs.microsoft.com/azure/active-directory/develop/active-directory-authentication-scenarios#daemon-or-server-application-to-web-api)), чтобы получить маркер доступа для Microsoft Graph в обмен на маркер доступа к начальной загрузке.
1. Azure AD возвращает в надстройку маркер доступа для Microsoft Graph (и маркер обновления, если надстройка запрашивает разрешение *offline_access*).
1. Серверный код кэширует маркер доступа в Microsoft Graph.
1. Серверный код отправляет запросы в Microsoft Graph и включает маркер доступа в Microsoft Graph.
1. Microsoft Graph возвращает данные надстройке, а она может передать их своему пользовательскому интерфейсу.
1. Когда маркер доступа в Microsoft Graph истекает, серверный код может использовать свой маркер обновления, чтобы получить новый маркер доступа в Microsoft Graph.

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>Разработка надстройки единого входа для Microsoft Graph

Надстройка, выполняющая вход Microsoft Graph, разрабатывается так же, как и любая другая надстройка с единым входом. Подробное описание см. в разделе [Включение единого входа для надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins). Разница заключается в том, что надстройка должна обязательно иметь веб-API на стороне сервера, а также в использовании термина "маркер доступа к начальной загрузке" вместо термина "маркер доступа" из указанной статьи. 

В зависимости от вашего языка и инфраструктуры могут быть доступны библиотеки, которые упростят код на стороне сервера, который вы должны написать. Ваш код должен выполнить следующее:

* Проверить в настройке маркер доступа к начальной загрузке, полученный от созданного ранее обработчика маркеров. Подробнее см. в статье [Проверка маркера доступа](sso-in-office-add-ins.md#validate-the-access-token). 
* Запустить поток "от имени" путем вызова конечной точки Azure AD версии 2.0, в ходе которого передается маркер доступа к надстройке, некоторые метаданные пользователя и учетные данные надстройки (идентификатор и секрет).
* Поместить в кэш возвращенный маркер доступа для Microsoft Graph. Подробнее об этом потоке см. в статье [Azure Active Directory v2.0 и поток "от имени" OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).
* Создать один или несколько методов веб-API, которые получают данные Microsoft Graph, передавая кэшированный маркер доступа в Microsoft Graph.

> [!NOTE]
> Примеры декодированных маркеров доступа для Microsoft Graph, которые были получены потоком "от имени", см. в статье [Azure Active Directory v2.0 и поток "от имени" OAuth 2.0](https://docs.microsoft.com/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of).

Примеры подробных пошаговых инструкций и сценариев см. ниже:

* [Создание надстройки Office на платформе Node.js с использованием единого входа](create-sso-office-add-ins-nodejs.md)
* [Создание надстройки Office на платформе ASP.NET с использованием единого входа](create-sso-office-add-ins-aspnet.md)
* [Сценарий: реализация единого входа для службы в надстройке Outlook](https://docs.microsoft.com/outlook/add-ins/implement-sso-in-outlook-add-in)



