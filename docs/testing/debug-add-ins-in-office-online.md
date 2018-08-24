---
title: Отладка надстроек в Office Online
description: Сведения о том, как тестировать и отлаживать надстройки в Office Online.
ms.date: 03/14/2018
ms.openlocfilehash: fac57e136c07bf33dce62908ea2c12d8be806f7b
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925341"
---
# <a name="debug-add-ins-in-office-online"></a><span data-ttu-id="7b796-103">Отладка надстроек в Office Online</span><span class="sxs-lookup"><span data-stu-id="7b796-103">Debug add-ins in Office Online</span></span>


<span data-ttu-id="7b796-p101">Вы можете создавать надстройки и выполнять их отладку на компьютере, на котором нет Windows или классического клиента Office 2013 или Office 2016 (например, если вы создаете надстройку на компьютере Mac). В этой статье рассказывается, как использовать Office Online для тестирования и отладки надстроек.</span><span class="sxs-lookup"><span data-stu-id="7b796-p101">You can build and debug add-ins on a computer that isn't running Windows or the Office 2013 or Office 2016 desktop client - for example, if you're developing on a Mac. This article describes how to use Office Online to test and debug your add-ins.</span></span> 

## <a name="prerequisites"></a><span data-ttu-id="7b796-106">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="7b796-106">Prerequisites</span></span>

<span data-ttu-id="7b796-107">Чтобы приступить к работе, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="7b796-107">To get started:</span></span>

- <span data-ttu-id="7b796-108">Получите учетную запись разработчика приложений для Office 365 (если у вас еще нет ее) или доступ к сайту SharePoint.</span><span class="sxs-lookup"><span data-stu-id="7b796-108">Get an Office 365 developer account if you don't already have one or have access to a SharePoint site.</span></span>
    
  > [!NOTE]
  > <span data-ttu-id="7b796-109">Чтобы бесплатно получить подписку разработчика приложений для Office 365, примите участие в нашей [программе для разработчиков приложений Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="7b796-109">To sign up for a free Office 365 developer subscription, join our [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span> <span data-ttu-id="7b796-110">Пошаговые инструкции для принятия участия в этой программе, регистрации и настройки подписки см. в [документации по программе для разработчиков приложений для Office 365](https://docs.microsoft.com/office/developer-program/office-365-developer-program).</span><span class="sxs-lookup"><span data-stu-id="7b796-110">See the [Office 365 Developer Program documentation](https://docs.microsoft.com/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription.</span></span>
     
- <span data-ttu-id="7b796-p103">Настройте каталог надстроек в Office 365 (SharePoint Online). Каталог надстроек — это специальное семейство веб-сайтов в SharePoint Online, в котором размещены библиотеки документов для надстроек Office. Если у вас есть сайт SharePoint, вы можете настроить библиотеку документов каталога надстроек. Дополнительные сведения см. в статье [Публикация надстроек области задач и контентных надстроек в каталоге надстроек в SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span><span class="sxs-lookup"><span data-stu-id="7b796-p103">Set up an add-in catalog on Office 365 (SharePoint Online). An add-in catalog is a dedicated site collection in SharePoint Online that hosts document libraries for Office Add-ins. If you have your own SharePoint site, you can set up an add-in catalog document library. For more information, see [Publish task pane and content add-ins to an add-in catalog on SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).</span></span>
    

## <a name="debug-your-add-in-from-excel-online-or-word-online"></a><span data-ttu-id="7b796-114">Отладка надстройки в Excel Online и Word Online</span><span class="sxs-lookup"><span data-stu-id="7b796-114">Debug your add-in from Excel Online or Word Online</span></span>

<span data-ttu-id="7b796-115">Для отладки надстройки с помощью Office Online выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="7b796-115">To debug your add-in by using Office Online:</span></span>

1. <span data-ttu-id="7b796-116">Разверните надстройку на сервере, поддерживающем SSL.</span><span class="sxs-lookup"><span data-stu-id="7b796-116">Deploy your add-in to a server that supports SSL.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="7b796-117">Рекомендуем использовать [генератор Yeoman](https://github.com/OfficeDev/generator-office) для создания и размещения надстройки.</span><span class="sxs-lookup"><span data-stu-id="7b796-117">We recommend that you use the [Yeoman generator](https://github.com/OfficeDev/generator-office) to create and host your add-in.</span></span>
     
2. <span data-ttu-id="7b796-p104">В [файле манифеста надстройки](../develop/add-in-manifests.md) измените значение элемента **SourceLocation** так, чтобы оно включало абсолютный URL-адрес, а не относительный. Пример:</span><span class="sxs-lookup"><span data-stu-id="7b796-p104">In your [add-in manifest file](../develop/add-in-manifests.md), update the **SourceLocation** element value to include an absolute, rather than a relative, URI. For example:</span></span>
      
    ```xml
    <SourceLocation DefaultValue="https://localhost:44300/App/Home/Home.html" />
    ```
    
3. <span data-ttu-id="7b796-120">Выложите манифест в библиотеку надстроек Office в каталоге надстроек в SharePoint.</span><span class="sxs-lookup"><span data-stu-id="7b796-120">Upload the manifest to the Office Add-ins library in the add-in catalog on SharePoint.</span></span>
    
4. <span data-ttu-id="7b796-121">В Office 365 в средстве запуска приложений запустите Excel Online или Word Online и откройте новый документ.</span><span class="sxs-lookup"><span data-stu-id="7b796-121">Launch Excel Online or Word Online from the app launcher in Office 365, and open a new document.</span></span>
    
5. <span data-ttu-id="7b796-122">Чтобы добавить вашу надстройку и протестировать ее в приложении, на вкладке "Вставка" выберите **Мои надстройки** или **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="7b796-122">On the Insert tab, choose  **My Add-ins** or **Office Add-ins** to insert your add-in and test it in the app.</span></span>
    
6. <span data-ttu-id="7b796-123">Выполните отладку надстройки в удобном для вас браузерном отладчике.</span><span class="sxs-lookup"><span data-stu-id="7b796-123">Use your favorite browser tool debugger to debug your add-in.</span></span>

## <a name="potential-issues"></a><span data-ttu-id="7b796-124">Возможные проблемы</span><span class="sxs-lookup"><span data-stu-id="7b796-124">Potential issues</span></span>    

<span data-ttu-id="7b796-125">Ниже указаны некоторые проблемы, которые могут возникнуть при отладке.</span><span class="sxs-lookup"><span data-stu-id="7b796-125">The following are some issues that you might encounter as you debug:</span></span>
    
- <span data-ttu-id="7b796-126">Причиной некоторых отображаемых ошибок JavaScript может быть Office Online.</span><span class="sxs-lookup"><span data-stu-id="7b796-126">Some JavaScript errors that you see might originate from Office Online.</span></span>
      
- <span data-ttu-id="7b796-127">Браузер может отобразить сообщение об ошибке, связанной с недопустимым сертификатом, которое необходимо обойти.</span><span class="sxs-lookup"><span data-stu-id="7b796-127">The browser might show an invalid certificate error that you will need to bypass.</span></span>
      
- <span data-ttu-id="7b796-128">Если вы задаете точки останова в коде, Office Online может отобразить сообщение об ошибке, свидетельствующее о том, что не удается выполнить сохранение.</span><span class="sxs-lookup"><span data-stu-id="7b796-128">If you set breakpoints in your code, Office Online might throw an error indicating that it is unable to save.</span></span>

## <a name="see-also"></a><span data-ttu-id="7b796-129">См. также</span><span class="sxs-lookup"><span data-stu-id="7b796-129">See also</span></span>

- [<span data-ttu-id="7b796-130">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7b796-130">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- <span data-ttu-id="7b796-131">
  [Политики проверки AppSource](https://docs.microsoft.com/office/dev/store/validation-policies)</span><span class="sxs-lookup"><span data-stu-id="7b796-131">[AppSource validation policies](https://docs.microsoft.com/office/dev/store/validation-policies)</span></span>  
- <span data-ttu-id="7b796-132">
  [Создание эффективных приложений и надстроек AppSource](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)</span><span class="sxs-lookup"><span data-stu-id="7b796-132">[Create effective AppSource apps and add-ins](https://docs.microsoft.com/office/dev/store/create-effective-office-store-listings)</span></span>  
- [<span data-ttu-id="7b796-133">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="7b796-133">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
    
