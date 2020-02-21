---
title: Требования к надстройкам Outlook
description: Чтобы надстройки Outlook загружались и работали надлежащим образом, существует ряд требований к серверам и клиентам.
ms.date: 10/09/2019
localization_priority: Priority
ms.openlocfilehash: 67aebd1fae19811797c07d33a5f6cac8907550f9
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166797"
---
# <a name="outlook-add-in-requirements"></a><span data-ttu-id="8fc86-103">Требования к надстройкам Outlook</span><span class="sxs-lookup"><span data-stu-id="8fc86-103">Outlook add-in requirements</span></span>

<span data-ttu-id="8fc86-104">Чтобы надстройки Outlook загружались и работали надлежащим образом, существует ряд требований к серверам и клиентам.</span><span class="sxs-lookup"><span data-stu-id="8fc86-104">For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.</span></span>

## <a name="client-requirements"></a><span data-ttu-id="8fc86-105">Требования к клиентам</span><span class="sxs-lookup"><span data-stu-id="8fc86-105">Client requirements</span></span>

- <span data-ttu-id="8fc86-106">Клиент должен быть одним из поддерживаемых ведущих приложений для надстроек Outlook. Эти клиенты поддерживают надстройки:</span><span class="sxs-lookup"><span data-stu-id="8fc86-106">The client must be one of the supported hosts for Outlook add-ins. The following clients support add-ins:</span></span>

   - <span data-ttu-id="8fc86-107">Outlook 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="8fc86-107">Outlook 2013 or later on Windows</span></span>
   - <span data-ttu-id="8fc86-108">Outlook 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="8fc86-108">Outlook 2016 or later on Mac</span></span>
   - <span data-ttu-id="8fc86-109">Outlook для iOS</span><span class="sxs-lookup"><span data-stu-id="8fc86-109">Outlook on iOS</span></span>
   - <span data-ttu-id="8fc86-110">Outlook для Android</span><span class="sxs-lookup"><span data-stu-id="8fc86-110">Outlook on Android</span></span>
   - <span data-ttu-id="8fc86-111">Outlook в Интернете для Exchange 2016 или более поздней версии и Office 365</span><span class="sxs-lookup"><span data-stu-id="8fc86-111">Outlook on the web for Exchange 2016 or later and Office 365</span></span>
   - <span data-ttu-id="8fc86-112">Outlook в Интернете для Exchange 2013</span><span class="sxs-lookup"><span data-stu-id="8fc86-112">Outlook on the web for Exchange 2013</span></span>
   - <span data-ttu-id="8fc86-113">Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="8fc86-113">Outlook.com</span></span>

- <span data-ttu-id="8fc86-p101">Клиент должен иметь прямое подключение к серверу Exchange Server или Office 365. При настройке клиента пользователь должен выбрать тип учетной записи **Exchange**, **Office 365** или **Outlook.com**. Если клиент настроен на подключение POP3 или IMAP, надстройки не загрузятся.</span><span class="sxs-lookup"><span data-stu-id="8fc86-p101">The client must be connected to an Exchange server or Office 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span></span>

## <a name="mail-server-requirements"></a><span data-ttu-id="8fc86-117">Требования к почтовым серверам</span><span class="sxs-lookup"><span data-stu-id="8fc86-117">Mail server requirements</span></span>

<span data-ttu-id="8fc86-p102">Если пользователь подключен к Office 365 или Outlook.com, требования к почтовому серверу уже выполнены. Но если пользователи подключаются к локально установленным экземплярам Exchange Server, требуется соответствие указанным ниже условиям.</span><span class="sxs-lookup"><span data-stu-id="8fc86-p102">If the user is connected to Office 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span></span>

- <span data-ttu-id="8fc86-120">Должен использоваться сервер Exchange 2013 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="8fc86-120">The server must be Exchange 2013 or later.</span></span>
- <span data-ttu-id="8fc86-121">Веб-службы Exchange (EWS) должны быть включены и подключены к Интернету.</span><span class="sxs-lookup"><span data-stu-id="8fc86-121">Exchange Web Services (EWS) must be enabled and must be exposed to the Internet.</span></span> <span data-ttu-id="8fc86-122">Многие надстройки требуют надлежащей работы EWS.</span><span class="sxs-lookup"><span data-stu-id="8fc86-122">Many add-ins require EWS to function properly.</span></span>
- <span data-ttu-id="8fc86-123">Чтобы сервер мог издавать действительные маркеры идентификации, он должен иметь действительный сертификат проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="8fc86-123">The server must have a valid authentication certificate in order for the server to issue valid identity tokens.</span></span> <span data-ttu-id="8fc86-124">Новые установленные экземпляры сервера Exchange Server обладают сертификатом проверки подлинности по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="8fc86-124">New installations of Exchange Server include a default authentication certificate.</span></span> <span data-ttu-id="8fc86-125">Дополнительные сведения см. в статьях [Цифровые сертификаты и шифрование в Exchange 2016](/Exchange/architecture/client-access/certificates) и [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span><span class="sxs-lookup"><span data-stu-id="8fc86-125">For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span></span>
- <span data-ttu-id="8fc86-126">Для получения доступа к надстройкам из [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) серверы клиентского доступа должны быть настроены на связь с AppSource.</span><span class="sxs-lookup"><span data-stu-id="8fc86-126">To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.</span></span>

## <a name="add-in-server-requirements"></a><span data-ttu-id="8fc86-127">Требования к серверам надстроек</span><span class="sxs-lookup"><span data-stu-id="8fc86-127">Add-in server requirements</span></span>

<span data-ttu-id="8fc86-p105">Файлы надстройки (например, HTML, JavaScript) могут быть размещены на любой платформе веб-сервера. Единственное требование — настройка сервера на использование HTTPS и доверия к SSL-сертификату со стороны клиента.</span><span class="sxs-lookup"><span data-stu-id="8fc86-p105">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span></span>

## <a name="see-also"></a><span data-ttu-id="8fc86-130">См. также</span><span class="sxs-lookup"><span data-stu-id="8fc86-130">See also</span></span>

- [<span data-ttu-id="8fc86-131">Требования для запуска надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8fc86-131">Requirements for running Office Add-ins</span></span>](../concepts/requirements-for-running-office-add-ins.md)
- [<span data-ttu-id="8fc86-132">Доступность ведущих приложений и платформ для надстроек Office (раздел Outlook)</span><span class="sxs-lookup"><span data-stu-id="8fc86-132">Office Add-in host and platform availability (Outlook section)</span></span>](../overview/office-add-in-availability.md#outlook)
- [<span data-ttu-id="8fc86-133">Поддержка наборов обязательных элементов API JavaScript для Outlook</span><span class="sxs-lookup"><span data-stu-id="8fc86-133">Outlook JavaScript API requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
