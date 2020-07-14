---
title: Требования к надстройкам Outlook
description: Чтобы надстройки Outlook загружались и работали надлежащим образом, существует ряд требований к серверам и клиентам.
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 700e0efd2ab2655de61d37d42038fa2c15a99cb4
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093996"
---
# <a name="outlook-add-in-requirements"></a><span data-ttu-id="0e47c-103">Требования к надстройкам Outlook</span><span class="sxs-lookup"><span data-stu-id="0e47c-103">Outlook add-in requirements</span></span>

<span data-ttu-id="0e47c-104">Чтобы надстройки Outlook загружались и работали надлежащим образом, существует ряд требований к серверам и клиентам.</span><span class="sxs-lookup"><span data-stu-id="0e47c-104">For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.</span></span>

## <a name="client-requirements"></a><span data-ttu-id="0e47c-105">Требования к клиентам</span><span class="sxs-lookup"><span data-stu-id="0e47c-105">Client requirements</span></span>

- <span data-ttu-id="0e47c-106">Клиент должен быть одним из поддерживаемых ведущих приложений для надстроек Outlook. Эти клиенты поддерживают надстройки:</span><span class="sxs-lookup"><span data-stu-id="0e47c-106">The client must be one of the supported hosts for Outlook add-ins. The following clients support add-ins:</span></span>

   - <span data-ttu-id="0e47c-107">Outlook 2013 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="0e47c-107">Outlook 2013 or later on Windows</span></span>
   - <span data-ttu-id="0e47c-108">Outlook 2016 или более поздней версии для Mac</span><span class="sxs-lookup"><span data-stu-id="0e47c-108">Outlook 2016 or later on Mac</span></span>
   - <span data-ttu-id="0e47c-109">Outlook для iOS</span><span class="sxs-lookup"><span data-stu-id="0e47c-109">Outlook on iOS</span></span>
   - <span data-ttu-id="0e47c-110">Outlook для Android</span><span class="sxs-lookup"><span data-stu-id="0e47c-110">Outlook on Android</span></span>
   - <span data-ttu-id="0e47c-111">Outlook в Интернете для Exchange 2016 или более поздней версии и Office 365</span><span class="sxs-lookup"><span data-stu-id="0e47c-111">Outlook on the web for Exchange 2016 or later and Office 365</span></span>
   - <span data-ttu-id="0e47c-112">Outlook в Интернете для Exchange 2013</span><span class="sxs-lookup"><span data-stu-id="0e47c-112">Outlook on the web for Exchange 2013</span></span>
   - <span data-ttu-id="0e47c-113">Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="0e47c-113">Outlook.com</span></span>

- <span data-ttu-id="0e47c-114">The client must be connected to an Exchange server or Microsoft 365 using a direct connection.</span><span class="sxs-lookup"><span data-stu-id="0e47c-114">The client must be connected to an Exchange server or Microsoft 365 using a direct connection.</span></span> <span data-ttu-id="0e47c-115">When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type.</span><span class="sxs-lookup"><span data-stu-id="0e47c-115">When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type.</span></span> <span data-ttu-id="0e47c-116">If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span><span class="sxs-lookup"><span data-stu-id="0e47c-116">If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span></span>

## <a name="mail-server-requirements"></a><span data-ttu-id="0e47c-117">Требования к почтовым серверам</span><span class="sxs-lookup"><span data-stu-id="0e47c-117">Mail server requirements</span></span>

<span data-ttu-id="0e47c-118">If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already.</span><span class="sxs-lookup"><span data-stu-id="0e47c-118">If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already.</span></span> <span data-ttu-id="0e47c-119">However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span><span class="sxs-lookup"><span data-stu-id="0e47c-119">However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span></span>

- <span data-ttu-id="0e47c-120">Должен использоваться сервер Exchange 2013 или более поздней версии.</span><span class="sxs-lookup"><span data-stu-id="0e47c-120">The server must be Exchange 2013 or later.</span></span>
- <span data-ttu-id="0e47c-121">Веб-службы Exchange (EWS) должны быть включены и подключены к Интернету.</span><span class="sxs-lookup"><span data-stu-id="0e47c-121">Exchange Web Services (EWS) must be enabled and must be exposed to the Internet.</span></span> <span data-ttu-id="0e47c-122">Многие надстройки требуют надлежащей работы EWS.</span><span class="sxs-lookup"><span data-stu-id="0e47c-122">Many add-ins require EWS to function properly.</span></span>
- <span data-ttu-id="0e47c-123">Чтобы сервер мог издавать действительные маркеры идентификации, он должен иметь действительный сертификат проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="0e47c-123">The server must have a valid authentication certificate in order for the server to issue valid identity tokens.</span></span> <span data-ttu-id="0e47c-124">Новые установленные экземпляры сервера Exchange Server обладают сертификатом проверки подлинности по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="0e47c-124">New installations of Exchange Server include a default authentication certificate.</span></span> <span data-ttu-id="0e47c-125">Дополнительные сведения см. в статьях [Цифровые сертификаты и шифрование в Exchange 2016](/Exchange/architecture/client-access/certificates) и [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span><span class="sxs-lookup"><span data-stu-id="0e47c-125">For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span></span>
- <span data-ttu-id="0e47c-126">Для получения доступа к надстройкам из [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) серверы клиентского доступа должны быть настроены на связь с AppSource.</span><span class="sxs-lookup"><span data-stu-id="0e47c-126">To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.</span></span>

## <a name="add-in-server-requirements"></a><span data-ttu-id="0e47c-127">Требования к серверам надстроек</span><span class="sxs-lookup"><span data-stu-id="0e47c-127">Add-in server requirements</span></span>

<span data-ttu-id="0e47c-128">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired.</span><span class="sxs-lookup"><span data-stu-id="0e47c-128">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired.</span></span> <span data-ttu-id="0e47c-129">The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span><span class="sxs-lookup"><span data-stu-id="0e47c-129">The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span></span>

## <a name="see-also"></a><span data-ttu-id="0e47c-130">См. также</span><span class="sxs-lookup"><span data-stu-id="0e47c-130">See also</span></span>

- [<span data-ttu-id="0e47c-131">Требования для запуска надстроек Office</span><span class="sxs-lookup"><span data-stu-id="0e47c-131">Requirements for running Office Add-ins</span></span>](../concepts/requirements-for-running-office-add-ins.md)
- [<span data-ttu-id="0e47c-132">Доступность ведущих приложений и платформ для надстроек Office (раздел Outlook)</span><span class="sxs-lookup"><span data-stu-id="0e47c-132">Office Add-in host and platform availability (Outlook section)</span></span>](../overview/office-add-in-availability.md#outlook)
- [<span data-ttu-id="0e47c-133">Поддержка наборов обязательных элементов API JavaScript для Outlook</span><span class="sxs-lookup"><span data-stu-id="0e47c-133">Outlook JavaScript API requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
