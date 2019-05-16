---
title: Веб-средства просмотра, используемые надстройками Office
description: ''
ms.date: 05/03/2019
localization_priority: Priority
ms.openlocfilehash: 6cb0d6e97dd559727b6a1e140d8417e1146e479a
ms.sourcegitcommit: 944cbb5c6ce055f6db1833182b24d490d1dce01d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/14/2019
ms.locfileid: "33992128"
---
# <a name="web-viewers-used-by-office-add-ins"></a><span data-ttu-id="87cd8-102">Веб-средства просмотра, используемые надстройками Office</span><span class="sxs-lookup"><span data-stu-id="87cd8-102">Web viewers used by Office Add-ins</span></span>

<span data-ttu-id="87cd8-103">Надстройки Office — это веб-приложения, поэтому им нужны веб-средство просмотра для отображения HTML-страниц веб-приложения и обработчик JavaScript для выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="87cd8-103">Since Office Add-ins are web applications, they need a web page viewer to display the HTML pages of the web application and a JavaScript engine to run the JavaScript.</span></span> <span data-ttu-id="87cd8-104">Оба поставляются браузером, установленным на компьютере пользователя.</span><span class="sxs-lookup"><span data-stu-id="87cd8-104">Both are supplied by a browser installed on the user’s computer.</span></span>

<span data-ttu-id="87cd8-105">Используемый браузер зависит от указанных ниже факторов.</span><span class="sxs-lookup"><span data-stu-id="87cd8-105">Which browser is used depends on:</span></span>

- <span data-ttu-id="87cd8-106">Операционная система компьютера.</span><span class="sxs-lookup"><span data-stu-id="87cd8-106">The computer’s operating system.</span></span>
- <span data-ttu-id="87cd8-107">Работает надстройка в Office Online, Office 365 или же Office 2013 либо более поздней версии без подписки.</span><span class="sxs-lookup"><span data-stu-id="87cd8-107">Whether the add-in is running in Office Online, Office 365, or non-subscription Office 2013 or later.</span></span>

<span data-ttu-id="87cd8-108">В приведенной ниже таблице указано, какой браузер используется для той или иной платформы и операционной системы.</span><span class="sxs-lookup"><span data-stu-id="87cd8-108">The following table shows which browser is used for the various platforms and operating systems.</span></span>

|<span data-ttu-id="87cd8-109">**ОС / платформа**</span><span class="sxs-lookup"><span data-stu-id="87cd8-109">**OS / Platform**</span></span>|<span data-ttu-id="87cd8-110">**Браузер**</span><span class="sxs-lookup"><span data-stu-id="87cd8-110">**Browser**</span></span>|
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
|<span data-ttu-id="87cd8-111">Office Online</span><span class="sxs-lookup"><span data-stu-id="87cd8-111">Office Online</span></span>|<span data-ttu-id="87cd8-112">Браузер, в котором открыт Office Online.</span><span class="sxs-lookup"><span data-stu-id="87cd8-112">The browser in which Office Online is opened.</span></span>|
|<span data-ttu-id="87cd8-113">Mac</span><span class="sxs-lookup"><span data-stu-id="87cd8-113">Mac</span></span>|<span data-ttu-id="87cd8-114">Safari</span><span class="sxs-lookup"><span data-stu-id="87cd8-114">Safari</span></span>|
|<span data-ttu-id="87cd8-115">iOS</span><span class="sxs-lookup"><span data-stu-id="87cd8-115">iOS</span></span>|<span data-ttu-id="87cd8-116">Safari</span><span class="sxs-lookup"><span data-stu-id="87cd8-116">Safari</span></span>|
|<span data-ttu-id="87cd8-117">Android</span><span class="sxs-lookup"><span data-stu-id="87cd8-117">Android</span></span>|<span data-ttu-id="87cd8-118">Chrome</span><span class="sxs-lookup"><span data-stu-id="87cd8-118">Chrome</span></span>|
|<span data-ttu-id="87cd8-119">Windows / Office 2013 или более поздней версии без подписки</span><span class="sxs-lookup"><span data-stu-id="87cd8-119">Windows / non-subscription Office 2013 or later</span></span>|<span data-ttu-id="87cd8-120">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="87cd8-120">Internet Explorer 11</span></span>|
|<span data-ttu-id="87cd8-121">Windows 10 версии</span><span class="sxs-lookup"><span data-stu-id="87cd8-121">Windows 10 ver.</span></span> <span data-ttu-id="87cd8-122">ниже 1903 / Office 365</span><span class="sxs-lookup"><span data-stu-id="87cd8-122">< 1903 / Office 365</span></span>|<span data-ttu-id="87cd8-123">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="87cd8-123">Internet Explorer 11</span></span>|
|<span data-ttu-id="87cd8-124">Windows 10 версии</span><span class="sxs-lookup"><span data-stu-id="87cd8-124">Windows 10 ver.</span></span> <span data-ttu-id="87cd8-125">1903 или выше / Office 365 версии ниже 16.0.11629</span><span class="sxs-lookup"><span data-stu-id="87cd8-125">>= 1903 / Office 365 ver < 16.0.11629</span></span>|<span data-ttu-id="87cd8-126">Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="87cd8-126">Internet Explorer 11</span></span>|
|<span data-ttu-id="87cd8-127">Windows 10 версии</span><span class="sxs-lookup"><span data-stu-id="87cd8-127">Windows 10 ver.</span></span> <span data-ttu-id="87cd8-128">1903 или выше / Office 365 версии 16.0.11629 или выше</span><span class="sxs-lookup"><span data-stu-id="87cd8-128">>= 1903 / Office 365 ver >= 16.0.11629</span></span>|<span data-ttu-id="87cd8-129">Microsoft Edge\*</span><span class="sxs-lookup"><span data-stu-id="87cd8-129">Microsoft Edge\*</span></span>|

<span data-ttu-id="87cd8-130">\* Если используется Microsoft Edge, экранный диктор Windows 10 (его иногда называют "читатель экрана") считывает тег `<title>` на странице, которая открывается в области задач.</span><span class="sxs-lookup"><span data-stu-id="87cd8-130">\* When Edge is being used, the Windows 10 Narrator (sometimes called a "screen reader") reads the `<title>` tag in the page that opens in the task pane.</span></span> <span data-ttu-id="87cd8-131">Когда используется Internet Explorer 11, экранный диктор читает панель заголовка области задач, полученный от значения `<DisplayName>` в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="87cd8-131">When Internet Explorer 11 is being used, the Narrator reads the title bar of the task pane, which comes from the `<DisplayName>` value in the add-in's manifest.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="87cd8-132">Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5.</span><span class="sxs-lookup"><span data-stu-id="87cd8-132">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="87cd8-133">Если какой-либо пользователь вашей надстройки применяет платформы с Internet Explorer 11, для применения синтаксиса и возможностей ECMAScript 2015 или более поздних версий вам нужно либо транскомпилировать свой код JavaScript в ES5, либо использовать полизаполнение.</span><span class="sxs-lookup"><span data-stu-id="87cd8-133">If any of your add-in's users have platforms that use Internet Explorer 11, then to use the syntax and features of ECMAScript 2015 or later, you will need to either transpile your JavaScript to ES5 or use a polyfill.</span></span> <span data-ttu-id="87cd8-134">Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.</span><span class="sxs-lookup"><span data-stu-id="87cd8-134">Also, Internet Explorer 11 does not support some HTML 5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="87cd8-135">Пока они не станут общедоступными, вам нужно быть участником программы предварительной оценки Windows, чтобы получить Windows версии 1903 или более поздней, а также участником программы предварительной оценки Office, чтобы получить Office версии 16.0.11629 или более высокой.</span><span class="sxs-lookup"><span data-stu-id="87cd8-135">Until they are generally available, you need to be a Windows Insider to get a Windows version 1903 or greater, and you need to be an Office Insider to get Office version 16.0.11629 or greater.</span></span>
>
> <span data-ttu-id="87cd8-136">Как стать участником программы предварительной оценки Windows:</span><span class="sxs-lookup"><span data-stu-id="87cd8-136">To join Windows Insiders:</span></span>
> 
> 1. <span data-ttu-id="87cd8-137">Перейдите по ссылке [Программа предварительной оценки Windows](https://insider.windows.com) и щелкните ссылку, чтобы стать ее участником.</span><span class="sxs-lookup"><span data-stu-id="87cd8-137">Go to [Windows Insider](https://insider.windows.com) and click the link to join Windows Insiders.</span></span>
> 2. <span data-ttu-id="87cd8-138">Откроется страница с указаниями о том, как использовать настройки Windows для включения предварительных сборок Windows.</span><span class="sxs-lookup"><span data-stu-id="87cd8-138">You will be taken to a page with instructions about how to use Windows Settings to enable preview builds of Windows.</span></span> <span data-ttu-id="87cd8-139">Следуйте инструкциям.</span><span class="sxs-lookup"><span data-stu-id="87cd8-139">Follow the instructions.</span></span> <span data-ttu-id="87cd8-140">При выборе скорости обновлений указывайте самый быстрый вариант.</span><span class="sxs-lookup"><span data-stu-id="87cd8-140">When you select the pace of updates, choose the fastest option.</span></span>
>
> <span data-ttu-id="87cd8-141">Как стать участником программы предварительной оценки Office:</span><span class="sxs-lookup"><span data-stu-id="87cd8-141">To join Office Insiders:</span></span>
> 
> 1. <span data-ttu-id="87cd8-142">Перейдите по ссылке [Программа предварительной оценки Office](https://insider.office.com/join).</span><span class="sxs-lookup"><span data-stu-id="87cd8-142">Go to [Get started as an Office Insider](https://insider.office.com/join).</span></span>
> 2. <span data-ttu-id="87cd8-143">Выполните приведенные на странице инструкции, чтобы стать участником.</span><span class="sxs-lookup"><span data-stu-id="87cd8-143">Follow the instruction on that page to join.</span></span> <span data-ttu-id="87cd8-144">Когда появится запрос на указание канала, выберите "Участник программы предварительной оценки".</span><span class="sxs-lookup"><span data-stu-id="87cd8-144">When asked to specify a channel, select Insider.</span></span>

## <a name="see-also"></a><span data-ttu-id="87cd8-145">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="87cd8-145">See also</span></span>

- [<span data-ttu-id="87cd8-146">Требования для запуска надстроек Office</span><span class="sxs-lookup"><span data-stu-id="87cd8-146">Requirements for Running Office Add-ins</span></span>](requirements-for-running-office-add-ins.md)
