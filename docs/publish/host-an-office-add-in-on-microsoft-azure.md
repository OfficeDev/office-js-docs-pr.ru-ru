---
title: Размещение надстройки Office в Microsoft Azure | Документация Майкрософт
description: Сведения о развертывании веб-приложения надстройки в Azure и загрузке неопубликованной надстройки для тестирования в клиентском приложении Office.
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: ce1cea8078c1842f4ce8cc57b8702c30393d8be8
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386836"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="b439a-103">Размещение надстройки Office в Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="b439a-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="b439a-p101">Самая простая надстройка Office состоит из XML-файла манифеста и HTML-страницы. В XML-файле манифеста описаны характеристики надстройки, например ее имя, сведения о том, в каких клиентах Office можно запускать надстройку, а также URL-адрес HTML-страницы надстройки. HTML-страница содержится в веб-приложении, с которым пользователь взаимодействует, когда устанавливает и запускает надстройку в клиентском приложении Office. Вы можете разместить веб-приложение надстройки Office на любой платформе веб-хостинга, включая Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="b439a-108">В этой статье рассказывается, как развернуть веб-приложение надстройки в Azure и [загрузить неопубликованную надстройку](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) для тестирования в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="b439a-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="b439a-109">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="b439a-109">Prerequisites</span></span> 

1. <span data-ttu-id="b439a-110">Установите [Visual Studio 2017](https://www.visualstudio.com/downloads) и не забудьте включить рабочую нагрузку **Разработка для Azure**.</span><span class="sxs-lookup"><span data-stu-id="b439a-110">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b439a-111">Если Visual Studio 2017 уже установлен, убедитесь, что рабочая нагрузка **Разработка для Azure** установлена, [используя установщик Visual Studio](https://docs.microsoft.com/visualstudio/install/modify-visual-studio).</span><span class="sxs-lookup"><span data-stu-id="b439a-111">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="b439a-112">Установите Office.</span><span class="sxs-lookup"><span data-stu-id="b439a-112">Install Office.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="b439a-113">Если у вас еще нет Office, можете [оформить бесплатную пробную подписку на 1 месяц](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span><span class="sxs-lookup"><span data-stu-id="b439a-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/en-US/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3.  <span data-ttu-id="b439a-114">Подпишитесь на Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-114">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="b439a-115">Если у вас еще нет подписки на Azure, вы можете [получить ее в рамках своей подписки на Visual Studio](https://azure.microsoft.com/ru-RU/pricing/member-offers/visual-studio-subscriptions/) или [зарегистрировать бесплатную учетную запись](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="b439a-115">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](https://azure.microsoft.com/ru-RU/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="b439a-116">Шаг 1. Создание общей папки для размещения XML-файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="b439a-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="b439a-117">Откройте проводник на своем компьютере разработчика.</span><span class="sxs-lookup"><span data-stu-id="b439a-117">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="b439a-118">Щелкните диск C: правой кнопкой мыши и выберите пункты **Создать** > **Папку**.</span><span class="sxs-lookup"><span data-stu-id="b439a-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="b439a-119">Назовите новую папку AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="b439a-119">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="b439a-120">Щелкните папку AddinManifests правой кнопкой мыши и выберите пункты **Общий доступ** > **Конкретные пользователи...**.</span><span class="sxs-lookup"><span data-stu-id="b439a-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="b439a-121">В окне **Общий доступ к файлам** щелкните стрелку раскрывающегося списка и выберите **Все** > **Добавить** > **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="b439a-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="b439a-p102">В этом руководстве для хранения XML-файла манифеста надстройки используется общая локальная папка. Для решения практических задач вы можете [развернуть XML-файл манифеста в каталоге SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) или [опубликовать надстройку в AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="b439a-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="b439a-124">Шаг 2. Добавление общей папки в доверенный каталог надстроек</span><span class="sxs-lookup"><span data-stu-id="b439a-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="b439a-125">Запустите Word и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="b439a-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b439a-126">В этом примере используется Word, но вы можете использовать любое приложение Office, поддерживающее надстройки Office, например Excel, Outlook, PowerPoint или Project.</span><span class="sxs-lookup"><span data-stu-id="b439a-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>
    
2.  <span data-ttu-id="b439a-127">Щелкните **Файл** > **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="b439a-127">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="b439a-128">В диалоговом окне **Параметры Word** щелкните **Центр управления безопасностью**, а затем — **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="b439a-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="b439a-p103">В диалоговом окне **Trust Center** выберите **Доверенные каталоги надстроек**. Введите UNC-путь для файлового ресурса, который вы создали ранее, в виде **URL-адреса каталога** (например, \\\YourMachineName\AddinManifests), а затем щелкните **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="b439a-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="b439a-131">Установите флажок **Показывать в меню**.</span><span class="sxs-lookup"><span data-stu-id="b439a-131">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="b439a-132">Когда XML-файл манифеста надстройки хранится в доверенном каталоге веб-надстроек, надстройка отображается в разделе **Общая папка** в диалоговом окне **Надстройки Office** (**Вставка** > **Мои надстройки**).</span><span class="sxs-lookup"><span data-stu-id="b439a-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="b439a-133">Закройте Word.</span><span class="sxs-lookup"><span data-stu-id="b439a-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="b439a-134">Шаг 3. Создание веб-приложения в Azure</span><span class="sxs-lookup"><span data-stu-id="b439a-134">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="b439a-135">Создайте пустое веб-приложение в Azure, используя [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) или [портал Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span><span class="sxs-lookup"><span data-stu-id="b439a-135">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="b439a-136">Использование Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="b439a-136">Using Visual Studio 2017</span></span>

<span data-ttu-id="b439a-137">Чтобы создать веб-приложение с помощью Visual Studio 2017, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b439a-137">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="b439a-p104">В Visual Studio в меню **Вид** меню выберите **обозреватель серверов**. Щелкните правой кнопкой мыши **Azure** и выберите пункт **Подключиться к подписке Microsoft Azure**. Чтобы подключиться к своей подписке Azure, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b439a-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="b439a-141">В Visual Studio в **обозревателе серверов** разверните пункт **Azure**, щелкните правой кнопкой мыши **Служба приложений** и выберите пункт **Создать службу приложений**.</span><span class="sxs-lookup"><span data-stu-id="b439a-141">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="b439a-142">В диалоговом окне **Создание службы приложений** укажите перечисленные ниже сведения.</span><span class="sxs-lookup"><span data-stu-id="b439a-142">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="b439a-p105">Введите уникальное **имя веб-приложения** для своего сайта. Azure проверит уникальность имени сайта в домене azurewebsites.net.</span><span class="sxs-lookup"><span data-stu-id="b439a-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="b439a-145">Выберите **подписку**, которую необходимо использовать для создания сайта.</span><span class="sxs-lookup"><span data-stu-id="b439a-145">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="b439a-p106">Выберите **группу ресурсов** для своего сайта. Если вы создадите группу, вам потребуется присвоить ей имя.</span><span class="sxs-lookup"><span data-stu-id="b439a-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="b439a-p107">Выберите **план службы приложений**, который необходимо использовать для создания сайта. Если вы создадите план, вам потребуется присвоить ему имя.</span><span class="sxs-lookup"><span data-stu-id="b439a-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="b439a-150">Нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="b439a-150">Choose **Create**.</span></span>

    <span data-ttu-id="b439a-151">Новое веб-приложение появится в **обозревателе серверов** в разделе **Azure** >> **Служба приложений** >> (выбранная группа ресурсов).</span><span class="sxs-lookup"><span data-stu-id="b439a-151">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="b439a-p108">Щелкните правой кнопкой мыши новое веб-приложение и выберите пункт **Открыть в браузере**. Откроется браузер, и в нем отобразится веб-страница с сообщением "Ваша служба приложений создана".</span><span class="sxs-lookup"><span data-stu-id="b439a-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="b439a-154">В адресной строке браузера измените URL-адрес веб-приложения так, чтобы он начинался со слова HTTPS, и нажмите клавишу **ВВОД**, чтобы убедиться, что протокол HTTPS включен.</span><span class="sxs-lookup"><span data-stu-id="b439a-154">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="b439a-155">Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="b439a-155">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="b439a-156">Использование портала Azure</span><span class="sxs-lookup"><span data-stu-id="b439a-156">Using the Azure portal</span></span>

<span data-ttu-id="b439a-157">Чтобы создать веб-приложение с помощью портала Azure, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b439a-157">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="b439a-158">Войдите в систему на [портале Azure](https://portal.azure.com/), используя свои учетные данные Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-158">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="b439a-159">Щелкните **Создать** > **Интернет и мобильные устройства** > **Веб-приложение**.</span><span class="sxs-lookup"><span data-stu-id="b439a-159">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="b439a-160">В диалоговом окне **Создание веб-приложения** укажите перечисленные ниже сведения.</span><span class="sxs-lookup"><span data-stu-id="b439a-160">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="b439a-p109">Введите уникальное **имя приложения** для своего сайта. Azure проверит уникальность имени сайта в домене azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="b439a-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="b439a-163">Выберите **подписку**, которую необходимо использовать для создания сайта.</span><span class="sxs-lookup"><span data-stu-id="b439a-163">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="b439a-p110">Выберите **группу ресурсов** для своего сайта. Если вы создадите группу, вам потребуется присвоить ей имя.</span><span class="sxs-lookup"><span data-stu-id="b439a-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="b439a-166">Выберите **операционную систему** для своего сайта.</span><span class="sxs-lookup"><span data-stu-id="b439a-166">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="b439a-p111">Выберите **план службы приложений**, который необходимо использовать для создания этого сайта. Если вы создадите план, вам потребуется присвоить ему имя.</span><span class="sxs-lookup"><span data-stu-id="b439a-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="b439a-169">Нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="b439a-169">Choose **Create**.</span></span>

4. <span data-ttu-id="b439a-170">Щелкните **Уведомления** (значок с изображением колокольчика, расположенный у верхнего края окна портала Azure) и выберите уведомление **Развертывания успешно выполнены**. Откроется страница **обзора** на портале Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-170">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="b439a-171">После развертывания сайта уведомление **Выполняется развертывание** сменится уведомлением **Успешные развертывания**.</span><span class="sxs-lookup"><span data-stu-id="b439a-171">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="b439a-p112">В разделе **Основное** на странице **обзора** сайта на портале Azure выберите URL-адрес, отображаемый в поле **URL-адрес**. Откроется браузер, и в нем отобразится веб-страница с сообщением "Ваша служба приложений создана".</span><span class="sxs-lookup"><span data-stu-id="b439a-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="b439a-174">В адресной строке браузера измените URL-адрес веб-приложения так, чтобы он начинался со слова HTTPS, и нажмите клавишу **ВВОД**, чтобы убедиться, что протокол HTTPS включен.</span><span class="sxs-lookup"><span data-stu-id="b439a-174">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="b439a-175">Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="b439a-175">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="b439a-176">Шаг 4. Создание надстройки Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="b439a-176">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="b439a-177">Запустите Visual Studio от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="b439a-177">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="b439a-178">Щелкните **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="b439a-178">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="b439a-179">В разделе **Шаблоны** разверните пункт **Visual C#** (или **Visual Basic**), затем пункт **Office/SharePoint** и выберите пункт **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="b439a-179">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="b439a-180">Выберите пункт **Веб-надстройка Word**, а затем нажмите кнопку **OK**, чтобы принять параметры, используемые по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="b439a-180">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="b439a-181">Visual Studio создаст базовую надстройку Word, которую вы можете опубликовать в том виде, в котором она есть, не внося изменений в ее веб-проект.</span><span class="sxs-lookup"><span data-stu-id="b439a-181">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="b439a-182">Действие 5. Публикация веб-приложения надстройки Office в Azure</span><span class="sxs-lookup"><span data-stu-id="b439a-182">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="b439a-183">Не закрывая проект вашей надстройки в Visual Studio, разверните узел решения в **обозревателе решений**, чтобы отображались оба проекта для решения.</span><span class="sxs-lookup"><span data-stu-id="b439a-183">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="b439a-p113">Щелкните правой кнопкой мыши веб-проект и выберите пункт **Опубликовать**. Веб-проект содержит файлы веб-приложения надстройки Office, так что это именно тот проект, который вы публикуете в Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="b439a-186">На вкладке **Публикация** выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b439a-186">On the **Publish** tab:</span></span>

      - <span data-ttu-id="b439a-187">Выберите пункт **Служба приложений Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="b439a-187">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="b439a-188">Щелкните **Выбрать существующую**.</span><span class="sxs-lookup"><span data-stu-id="b439a-188">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="b439a-189">Щелкните **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="b439a-189">Choose **Publish**.</span></span> 

6. <span data-ttu-id="b439a-190">В диалоговом окне **Служба приложений** найдите и выберите веб-приложение, которое вы создали на [шаге 3](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="b439a-190">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="b439a-p114">Visual Studio опубликует веб-проект надстройки Office в вашем веб-приложении Azure. Когда Visual Studio завершит публикацию веб-проекта, откроется браузер, в котором отобразится веб-страница с текстом "Приложение службы приложений создано". Это текущая страница, используемая по умолчанию, для веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="b439a-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="b439a-194">Чтобы отобразить веб-страницу для вашей надстройки, измените URL-адрес так, чтобы в нем использовался протокол HTTPS и был указан путь к HTML-странице вашей надстройки (пример: https://YourDomain.azurewebsites.net/Home.html).</span><span class="sxs-lookup"><span data-stu-id="b439a-194">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="b439a-195">Это подтверждает, что веб-приложение вашей надстройки теперь размещено в Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-195">This confirms that your add-in's web app is now hosted on Azure.</span></span> <span data-ttu-id="b439a-196">Скопируйте URL-адрес корня (пример: https://YourDomain.azurewebsites.net); он потребуется, когда вы будете редактировать файл манифеста надстройки далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="b439a-196">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="b439a-197">Шаг 6. Редактирование и развертывание XML-файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="b439a-197">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="b439a-198">В Visual Studio (с примером надстройки Office, открытом в **обозревателе решений**) разверните решение так, чтобы отображались оба проекта.</span><span class="sxs-lookup"><span data-stu-id="b439a-198">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="b439a-p116">Разверните проект надстройки Office (например, WordWebAddIn), щелкните правой кнопкой мыши папку манифеста, а затем нажмите кнопку **Открыть**. Откроется XML-файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="b439a-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="b439a-201">В XML-файле манифеста найдите и замените все фрагменты ~remoteAppUrl URL-адресом корня веб-приложения надстройки в Azure.</span><span class="sxs-lookup"><span data-stu-id="b439a-201">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span></span> <span data-ttu-id="b439a-202">Это URL-адрес, который вы скопировали ранее после публикации веб-приложения надстройки в Azure (пример: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="b439a-202">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="b439a-p118">Щелкните **Файл** и выберите пункт **Сохранить все**. Закройте XML-файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="b439a-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="b439a-205">Вернитесь в **обозреватель решений**, щелкните правой кнопкой мыши папку манифеста и выберите пункт **Открыть папку в проводнике**.</span><span class="sxs-lookup"><span data-stu-id="b439a-205">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="b439a-206">Скопируйте XML-файл манифеста надстройки (например, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="b439a-206">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="b439a-207">Откройте сетевой файловый ресурс, который вы создали в [действии 1 "Создание общей папки"](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) и вставьте файл манифеста в папку.</span><span class="sxs-lookup"><span data-stu-id="b439a-207">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="b439a-208">Шаг 7. Вставка и запуск надстройки в клиентском приложении Office</span><span class="sxs-lookup"><span data-stu-id="b439a-208">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="b439a-209">Запустите Word и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="b439a-209">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="b439a-210">На ленте щелкните **Вставка** > **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="b439a-210">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="b439a-p119">В диалоговом окне **Надстройки Office** выберите **ОБЩАЯ ПАПКА**. Word выполнит сканирование папки, которую вы указали в качестве надежного каталога надстроек (в [действии 2 "Добавление файлового ресурса в надежный каталог надстроек"](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) и отобразит надстройки в диалоговом окне. Должен отобразиться значок для вашего примера надстройки.</span><span class="sxs-lookup"><span data-stu-id="b439a-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="b439a-p120">Щелкните значок своей надстройки и нажмите кнопку **Добавить**. На ленту будет добавлена кнопка **Показать область задач** для вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="b439a-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="b439a-p121">На ленте вкладки **Главная** нажмите кнопку **Показать область задач**. Надстройка откроется в области задач справа от текущего документа.</span><span class="sxs-lookup"><span data-stu-id="b439a-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="b439a-p122">Убедитесь, что надстройка работает, выбрав любой текст в документе и нажав кнопку **Highlight!** (Выделить!) в области задач.</span><span class="sxs-lookup"><span data-stu-id="b439a-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="b439a-220">См. также</span><span class="sxs-lookup"><span data-stu-id="b439a-220">See also</span></span>

- [<span data-ttu-id="b439a-221">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="b439a-221">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="b439a-222">Упаковка надстройки с помощью Visual Studio для публикации</span><span class="sxs-lookup"><span data-stu-id="b439a-222">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
