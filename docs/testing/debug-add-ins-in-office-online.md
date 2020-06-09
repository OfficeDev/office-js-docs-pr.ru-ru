---
title: Отладка надстроек в Office в Интернете
description: Сведения о том, как тестировать и отлаживать надстройки в Office в Интернете.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4f2aa498e5d9f49fcdf306ac2c4c80ea6fbd496c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611241"
---
# <a name="debug-add-ins-in-office-on-the-web"></a><span data-ttu-id="a5df1-103">Отладка надстроек в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="a5df1-103">Debug add-ins in Office on the web</span></span>


<span data-ttu-id="a5df1-104">Вы можете создавать надстройки и выполнять их отладку на компьютере, на котором нет Windows или классического клиента Office (например, если вы создаете надстройку на компьютере Mac).</span><span class="sxs-lookup"><span data-stu-id="a5df1-104">You can build and debug add-ins on a computer that isn't running Windows or the Office desktop client&mdash;for example, if you're developing on a Mac.</span></span> <span data-ttu-id="a5df1-105">В этой статье описано, как использовать Office а Интернете для тестирования и отладки надстроек.</span><span class="sxs-lookup"><span data-stu-id="a5df1-105">This article describes how to use Office on the web to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="a5df1-106">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="a5df1-106">Prerequisites</span></span>

<span data-ttu-id="a5df1-107">Чтобы приступить к работе, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="a5df1-107">To get started:</span></span>

- <span data-ttu-id="a5df1-108">Получите учетную запись разработчика приложений для Office 365 (если у вас еще нет ее) или доступ к сайту SharePoint.</span><span class="sxs-lookup"><span data-stu-id="a5df1-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>

  > [!NOTE]
  > <span data-ttu-id="a5df1-p102">Чтобы получить бесплатную возобновляемую подписку разработчика приложений для Office 365 на 90 дней, примите участие в нашей [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program). Пошаговые инструкции для присоединения к этой программе и настройки подписки см. в [документации по программе для разработчиков Office 365](/office/developer-program/office-365-developer-program).</span><span class="sxs-lookup"><span data-stu-id="a5df1-p102">To get a free, 90-day renewable Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program). See the [Office 365 Developer Program documentation](/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and configure your subscription.</span></span>

- <span data-ttu-id="a5df1-p103">Настройте каталог приложений в Office 365 (SharePoint Online). Каталог приложений — это специальное семейство веб-сайтов в SharePoint Online, в котором размещены библиотеки документов для надстроек Office. Если у вас есть сайт SharePoint, вы можете настроить библиотеку документов каталога приложений. Дополнительные сведения см. в статье [Публикация надстроек области задач и контентных надстроек в каталоге приложений SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span><span class="sxs-lookup"><span data-stu-id="a5df1-p103">Set up an app catalog on Office 365 (SharePoint Online). An app catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an app catalog document library. For more information, see [Publish task pane and content add-ins to an app catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>


## <a name="debug-your-add-in-from-excel-or-word-on-the-web"></a><span data-ttu-id="a5df1-114">Отладка надстройки в Excel и Word в Интернете</span><span class="sxs-lookup"><span data-stu-id="a5df1-114">Debug your add-in from Excel or Word on the web</span></span>

<span data-ttu-id="a5df1-115">Для отладки надстройки с помощью Office в Интернете выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="a5df1-115">To debug your add-in by using Office on the web:</span></span>

1. <span data-ttu-id="a5df1-116">Разверните надстройку на сервере, поддерживающем SSL.</span><span class="sxs-lookup"><span data-stu-id="a5df1-116">Deploy your add-in to a server that supports SSL.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a5df1-117">Рекомендуем использовать [генератор Yeoman](https://github.com/OfficeDev/generator-office) для создания и размещения надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5df1-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>

2. <span data-ttu-id="a5df1-p104">В [файле манифеста надстройки](../develop/add-in-manifests.md) измените значение элемента **SourceLocation** так, чтобы оно включало абсолютный URL-адрес, а не относительный. Пример:</span><span class="sxs-lookup"><span data-stu-id="a5df1-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>

    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```

3. <span data-ttu-id="a5df1-120">Отправьте манифест в библиотеку надстроек Office в каталоге приложений SharePoint.</span><span class="sxs-lookup"><span data-stu-id="a5df1-120">Upload the manifest to the Office Add-ins library in the app catalog on SharePoint.</span></span>

4. <span data-ttu-id="a5df1-121">В Office 365 с помощью средства запуска приложений запустите Excel или Word в Интернете и откройте новый документ.</span><span class="sxs-lookup"><span data-stu-id="a5df1-121">Launch Excel or Word on the web from the app launcher in Office 365, and open a new document.</span></span>

5. <span data-ttu-id="a5df1-122">Чтобы добавить вашу надстройку и протестировать ее в приложении, на вкладке "Вставка" выберите **Мои надстройки** или **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="a5df1-122">On the Insert tab, choose **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>

6. <span data-ttu-id="a5df1-123">Выполните отладку надстройки в удобном для вас браузерном отладчике.</span><span class="sxs-lookup"><span data-stu-id="a5df1-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="a5df1-124">Возможные проблемы</span><span class="sxs-lookup"><span data-stu-id="a5df1-124">Potential issues</span></span>

<span data-ttu-id="a5df1-125">Ниже указаны некоторые проблемы, которые могут возникнуть при отладке.</span><span class="sxs-lookup"><span data-stu-id="a5df1-125">The following are some issues that you might encounter as you debug:</span></span>

- <span data-ttu-id="a5df1-126">Причиной некоторых отображаемых ошибок JavaScript может быть Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a5df1-126">Some JavaScript errors that you see might originate from Office on the web.</span></span>

- <span data-ttu-id="a5df1-127">В браузере может появиться сообщение об ошибке, связанной с недопустимым сертификатом, которое необходимо обойти.</span><span class="sxs-lookup"><span data-stu-id="a5df1-127">The browser might show an invalid certificate error that you will need to bypass.</span></span> <span data-ttu-id="a5df1-128">Этот процесс зависит от браузера, и пользовательские интерфейсы различных браузеров, предназначенные для его выполнения, периодически изменяются.</span><span class="sxs-lookup"><span data-stu-id="a5df1-128">The process for doing this varies with the browser and the various browsers' UIs for doing this change periodically.</span></span> <span data-ttu-id="a5df1-129">Инструкции можно найти в справке браузера или выполнить поиск в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a5df1-129">You should search the browser's help or search online for instructions.</span></span> <span data-ttu-id="a5df1-130">(Например, выполните поиск по запросу "Предупреждение Microsoft Edge о недействительном сертификате".) В большинстве браузеров на странице предупреждения содержится ссылка, позволяющая перейти к странице надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5df1-130">(For example, search for "Microsoft Edge invalid certificate warning".) Most browsers will have a link on the warning page that enables you to click through to the add-in page.</span></span> <span data-ttu-id="a5df1-131">Например, в Microsoft Edge есть ссылка "Перейти на веб-страницу (не рекомендуется)".</span><span class="sxs-lookup"><span data-stu-id="a5df1-131">For example, Microsoft Edge has a link "Go on to the webpage (Not recommended)".</span></span> <span data-ttu-id="a5df1-132">При этом, как правило, вам потребуется переходить по этой ссылке при каждой перезагрузке надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5df1-132">But you will usually have to go through this link every time the add-in reloads.</span></span> <span data-ttu-id="a5df1-133">Сведения о более долговечных способах обхода см. в предложенных разделах справки.</span><span class="sxs-lookup"><span data-stu-id="a5df1-133">For a longer lasting bypass, see the help as suggested.</span></span>

- <span data-ttu-id="a5df1-134">Если вы задаете точки останова в коде, Office в Интернете может вывести сообщение об ошибке, свидетельствующее о том, что не удается выполнить сохранение.</span><span class="sxs-lookup"><span data-stu-id="a5df1-134">If you set breakpoints in your code, Office on the web might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="a5df1-135">См. также</span><span class="sxs-lookup"><span data-stu-id="a5df1-135">See also</span></span>

- [<span data-ttu-id="a5df1-136">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a5df1-136">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="a5df1-137">Политики проверки AppSource</span><span class="sxs-lookup"><span data-stu-id="a5df1-137">AppSource validation policies</span></span>](/legal/marketplace/certification-policies)  
- [<span data-ttu-id="a5df1-138">Создание эффективных приложений и надстроек AppSource</span><span class="sxs-lookup"><span data-stu-id="a5df1-138">Create effective AppSource apps and add-ins</span></span>](/office/dev/store/create-effective-office-store-listings)  
- [<span data-ttu-id="a5df1-139">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="a5df1-139">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
