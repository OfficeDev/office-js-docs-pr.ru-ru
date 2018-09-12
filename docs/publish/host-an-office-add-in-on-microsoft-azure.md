---
title: Размещение надстройки Office в Microsoft Azure
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 62fc3c6dc212efc47493f2bcb3a994fb4db6a752
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945567"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="581a8-102">Размещение надстройки Office в Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="581a8-102">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="581a8-p101">Самая простая надстройка Office состоит из XML-файла манифеста и HTML-страницы. В XML-файле манифеста описаны характеристики надстройки, например ее имя, сведения о том, в каких клиентах Office можно запускать надстройку, а также URL-адрес HTML-страницы надстройки. HTML-страница содержится в веб-приложении, с которым пользователь взаимодействует, когда устанавливает и запускает надстройку в клиентском приложении Office. Вы можете разместить веб-приложение надстройки Office на любой платформе веб-хостинга, включая Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-p101">The simplest Office Add-in is made up of an XML manifest file and an HTML page. The XML manifest file describes the add-in's characteristics, such as its name, what Office client applications it can run in, and the URL for the add-in's HTML page. The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application. You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="581a8-107">В этой статье описано, как развернуть веб-приложение надстройки в Azure и [загрузить неопубликованную надстройку](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) для тестирования в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="581a8-107">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="581a8-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="581a8-108">Prerequisites</span></span> 

1. <span data-ttu-id="581a8-109">Установите [Visual Studio 2017](https://www.visualstudio.com/downloads) и не забудьте включить рабочую нагрузку **Разработка для Azure**.</span><span class="sxs-lookup"><span data-stu-id="581a8-109">Install [Visual Studio 2017](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="581a8-110">Если Visual Studio 2017 уже установлен, убедитесь, что рабочая нагрузка **Разработка для Azure** установлена, [используя установщик Visual Studio](https://docs.microsoft.com/visualstudio/install/modify-visual-studio).</span><span class="sxs-lookup"><span data-stu-id="581a8-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="581a8-111">Установите Office.</span><span class="sxs-lookup"><span data-stu-id="581a8-111">Install Office.</span></span> 
    
    > [!NOTE]
    > <span data-ttu-id="581a8-112">Если у вас еще нет Office, можете [оформить бесплатную пробную подписку на 1 месяц](http://office.microsoft.com/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span><span class="sxs-lookup"><span data-stu-id="581a8-112">If you don't already have Office 2016, you can [register for a free 1-month trial](http://office.microsoft.com/try/?WT%2Eintid1=ODC%5FENUS%5FFX101785584%5FXT104056786).</span></span>

3.  <span data-ttu-id="581a8-113">Подпишитесь на Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-113">Obtain an Azure subscription.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="581a8-114">Если у вас еще нет подписки на Azure, вы можете [получить ее в рамках своей подписки на Visual Studio](http://www.windowsazure.com/pricing/member-offers/msdn-benefits/) или [зарегистрировать бесплатную учетную запись](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="581a8-114">If don't already have an Azure subscription, you can [get one as part of your MSDN subscription](http://www.windowsazure.com/pricing/member-offers/msdn-benefits/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="581a8-115">Шаг 1. Создание общей папки для размещения XML-файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="581a8-115">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="581a8-116">Откройте проводник на своем компьютере разработчика.</span><span class="sxs-lookup"><span data-stu-id="581a8-116">Open File Explorer on your development computer.</span></span>
    
2. <span data-ttu-id="581a8-117">Щелкните диск C: правой кнопкой мыши и выберите пункты **Создать** > **Папку**.</span><span class="sxs-lookup"><span data-stu-id="581a8-117">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>
    
3. <span data-ttu-id="581a8-118">Назовите новую папку AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="581a8-118">Name the new folder AddinManifests.</span></span>
    
4. <span data-ttu-id="581a8-119">Щелкните папку AddinManifests правой кнопкой мыши и выберите пункты **Общий доступ** > **Конкретные пользователи...**.</span><span class="sxs-lookup"><span data-stu-id="581a8-119">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>
    
5. <span data-ttu-id="581a8-120">В окне **Общий доступ к файлам** щелкните стрелку раскрывающегося списка и выберите **Все** > **Добавить** > **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="581a8-120">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>
    
> [!NOTE]
> <span data-ttu-id="581a8-p102">В этом руководстве для хранения XML-файла манифеста надстройки используется общая локальная папка. Для решения практических задач вы можете [развернуть XML-файл манифеста в каталоге SharePoint](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) или [опубликовать надстройку в AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span><span class="sxs-lookup"><span data-stu-id="581a8-p102">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file. In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="581a8-123">Шаг 2. Добавление общей папки в доверенный каталог надстроек</span><span class="sxs-lookup"><span data-stu-id="581a8-123">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1.  <span data-ttu-id="581a8-124">Запустите Word и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="581a8-124">Start Word 2016 and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="581a8-125">В этом примере используется Word, но вы можете использовать любое приложение Office, поддерживающее надстройки Office, например, Excel, Outlook, PowerPoint или Project.</span><span class="sxs-lookup"><span data-stu-id="581a8-125">Although this example uses Word 2016, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project 2016.</span></span>
    
2.  <span data-ttu-id="581a8-126">Щелкните **Файл**  >  **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="581a8-126">Choose **File** > **Options**.</span></span>
    
3.  <span data-ttu-id="581a8-127">В диалоговом окне **Параметры Word** щелкните **Центр управления безопасностью**, а затем — **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="581a8-127">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span> 
    
4.  <span data-ttu-id="581a8-p103">В диалоговом окне **Trust Center** выберите **Доверенные каталоги надстроек**. Введите UNC-путь для файлового ресурса, который вы создали ранее, в виде **URL-адреса каталога** (например, \\\YourMachineName\AddinManifests), а затем щелкните **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="581a8-p103">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**. Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 
    
5. <span data-ttu-id="581a8-130">Установите флажок **Показывать в меню**.</span><span class="sxs-lookup"><span data-stu-id="581a8-130">Select the check box for **Show in Menu**.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="581a8-131">Когда XML-файл манифеста надстройки хранится в доверенном каталоге веб-надстроек, надстройка отображается в разделе **Общая папка** в диалоговом окне **Надстройки Office** (**Вставка** > **Мои надстройки**).</span><span class="sxs-lookup"><span data-stu-id="581a8-131">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="581a8-132">Закройте Word.</span><span class="sxs-lookup"><span data-stu-id="581a8-132">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure"></a><span data-ttu-id="581a8-133">Шаг 3. Создание веб-приложение в Azure</span><span class="sxs-lookup"><span data-stu-id="581a8-133">Step 3: Create a web app in Azure</span></span>

<span data-ttu-id="581a8-134">Создайте пустое веб-приложение в Azure, используя [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) или [портал Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span><span class="sxs-lookup"><span data-stu-id="581a8-134">Create an empty web app in Azure either by using [Visual Studio 2017](../publish/host-an-office-add-in-on-microsoft-azure.md#using-visual-studio-2017) or by using the [Azure portal](../publish/host-an-office-add-in-on-microsoft-azure.md#using-the-azure-portal).</span></span>

### <a name="using-visual-studio-2017"></a><span data-ttu-id="581a8-135">Использование Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="581a8-135">Using Visual Studio 2017</span></span>

<span data-ttu-id="581a8-136">Чтобы создать веб-приложение с помощью Visual Studio 2017, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="581a8-136">To create the web app using Visual Studio 2017, complete the following steps.</span></span>

1. <span data-ttu-id="581a8-p104">В Visual Studio в меню **Вид** меню выберите **обозреватель серверов**. Щелкните правой кнопкой мыши **Azure** и выберите пункт **Подключиться к подписке Microsoft Azure**. Чтобы подключиться к своей подписке Azure, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="581a8-p104">In Visual Studio, in the **View** menu, choose **Server Explorer**. Right-click **Azure** and choose **Connect to Microsoft Azure subscription**. Follow the instructions for connecting to your Azure subscription.</span></span>
    
2. <span data-ttu-id="581a8-140">В Visual Studio в **обозревателе серверов** разверните пункт **Azure**, щелкните правой кнопкой мыши **Служба приложений** и выберите пункт **Создать службу приложений**.</span><span class="sxs-lookup"><span data-stu-id="581a8-140">In Visual Studio, in **Server Explorer**, expand **Azure**, right-click **App Service**, and then choose **Create New App Service**.</span></span>
    
3. <span data-ttu-id="581a8-141">В диалоговом окне **Создание службы приложений** укажите перечисленные ниже сведения.</span><span class="sxs-lookup"><span data-stu-id="581a8-141">In the **Create App Service** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="581a8-p105">Введите уникальное **имя веб-приложения** для своего сайта. Azure проверит уникальность имени сайта в домене azurewebsites.net.</span><span class="sxs-lookup"><span data-stu-id="581a8-p105">Enter a unique **Web App Name** for your site. Azure verifies that the site name is unique across the azurewebsites.net domain.</span></span>

      - <span data-ttu-id="581a8-144">Выберите **подписку**, которую необходимо использовать для создания сайта.</span><span class="sxs-lookup"><span data-stu-id="581a8-144">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="581a8-p106">Выберите **группу ресурсов** для своего сайта. Если вы создадите группу, вам потребуется присвоить ей имя.</span><span class="sxs-lookup"><span data-stu-id="581a8-p106">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>
    
      - <span data-ttu-id="581a8-p107">Выберите **план службы приложений**, который необходимо использовать для создания сайта. Если вы создадите план, вам потребуется присвоить ему имя.</span><span class="sxs-lookup"><span data-stu-id="581a8-p107">Choose the **App Service Plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="581a8-149">Нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="581a8-149">Choose **Create**.</span></span>

    <span data-ttu-id="581a8-150">Новое веб-приложение появится в **обозревателе серверов** в разделе **Azure** >> **Служба приложений** >> (выбранная группа ресурсов).</span><span class="sxs-lookup"><span data-stu-id="581a8-150">The new web app appears in **Server Explorer** under **Azure** >> **App Service** >> (the chosen resouce group).</span></span>
    
4. <span data-ttu-id="581a8-p108">Щелкните правой кнопкой мыши новое веб-приложение и выберите пункт **Открыть в браузере**. Откроется браузер, и в нем отобразится веб-страница с сообщением "Ваша служба приложений создана".</span><span class="sxs-lookup"><span data-stu-id="581a8-p108">Right-click the new web app and then choose **View in Browser**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span>
    
5. <span data-ttu-id="581a8-153">В адресной строке браузера измените URL-адрес веб-приложения так, чтобы он начинался со слова HTTPS, и нажмите клавишу **ВВОД**, чтобы убедиться, что протокол HTTPS включен.</span><span class="sxs-lookup"><span data-stu-id="581a8-153">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="581a8-154"> Сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="581a8-154">Azure websites automatically provide an HTTPS endpoint.</span></span>
    
### <a name="using-the-azure-portal"></a><span data-ttu-id="581a8-155">Использование портала Azure</span><span class="sxs-lookup"><span data-stu-id="581a8-155">Using the Azure portal</span></span>

<span data-ttu-id="581a8-156">Чтобы создать веб-приложение с помощью портала Azure, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="581a8-156">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="581a8-157">Войдите в систему на [портале Azure](https://portal.azure.com/), используя свои учетные данные Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-157">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>
    
2. <span data-ttu-id="581a8-158">Щелкните **Создать** > **Интернет и мобильные устройства** > **Веб-приложение**.</span><span class="sxs-lookup"><span data-stu-id="581a8-158">Choose **New** > **Web + Mobile** > **Web App**.</span></span> 

3. <span data-ttu-id="581a8-159">В диалоговом окне **Создание веб-приложения** укажите перечисленные ниже сведения.</span><span class="sxs-lookup"><span data-stu-id="581a8-159">In the **Web App Create** dialog box, provide this information:</span></span>
    
      - <span data-ttu-id="581a8-p109">Введите уникальное **имя приложения** для своего сайта. Azure проверит уникальность имени сайта в домене azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="581a8-p109">Enter a unique **App name** for your site. Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="581a8-162">Выберите **подписку**, которую необходимо использовать для создания сайта.</span><span class="sxs-lookup"><span data-stu-id="581a8-162">Choose the **Subscription** to use for creating this site.</span></span>

      - <span data-ttu-id="581a8-p110">Выберите **группу ресурсов** для своего сайта. Если вы создадите группу, вам потребуется присвоить ей имя.</span><span class="sxs-lookup"><span data-stu-id="581a8-p110">Choose the **Resource Group** for your site. If you create a new group, you also need to name it.</span></span>

      - <span data-ttu-id="581a8-165">Выберите **операционную систему** для своего сайта.</span><span class="sxs-lookup"><span data-stu-id="581a8-165">Choose the **OS** for your site.</span></span>
    
      - <span data-ttu-id="581a8-p111">Выберите **план службы приложений**, который необходимо использовать для создания этого сайта. Если вы создадите план, вам потребуется присвоить ему имя.</span><span class="sxs-lookup"><span data-stu-id="581a8-p111">Choose the **App Service plan** to use for creating this site. If you create a new plan, you also need to name it.</span></span>
       
      - <span data-ttu-id="581a8-168">Нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="581a8-168">Choose **Create**.</span></span>

4. <span data-ttu-id="581a8-169">Щелкните **Уведомления** (значок с изображением колокольчика, расположенный у верхнего края окна портала Azure) и выберите уведомление **Развертывания успешно выполнены**. Откроется страница **обзора** на портале Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-169">Choose **Notifications** (the bell icon that is located along the top edge of the Azure portal) and then choose the **Deployments succeeded** notification to open the site's **Overview** page in the Azure portal.</span></span>

    > [!NOTE]
    > <span data-ttu-id="581a8-170">После развертывания сайта уведомление **Выполняется развертывание** сменится уведомлением **Успешные развертывания**.</span><span class="sxs-lookup"><span data-stu-id="581a8-170">The notification will change from **Deployment in progress** to **Deployments succeeded** when the site deployment completes.</span></span>

5. <span data-ttu-id="581a8-p112">В разделе **Основное** на странице **обзора** сайта на портале Azure выберите URL-адрес, отображаемый в поле **URL-адрес**. Откроется браузер, и в нем отобразится веб-страница с сообщением "Ваша служба приложений создана".</span><span class="sxs-lookup"><span data-stu-id="581a8-p112">In the **Essentials** section of the site's **Overview** page in the Azure portal, choose the URL that is displayed under **URL**. Your browser opens and displays a webpage with the message "Your App Service app has been created."</span></span> 
    
6. <span data-ttu-id="581a8-173">В адресной строке браузера измените URL-адрес веб-приложения так, чтобы он начинался со слова HTTPS, и нажмите клавишу **ВВОД**, чтобы убедиться, что протокол HTTPS включен.</span><span class="sxs-lookup"><span data-stu-id="581a8-173">In the browser address bar, change the URL for the web app so that it uses HTTPS and press **Enter** to confirm that the HTTPS protocol is enabled.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="581a8-174"> Сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="581a8-174">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="581a8-175">Шаг 4. Создайте надстройку Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="581a8-175">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="581a8-176">Запустите Visual Studio от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="581a8-176">Start Visual Studio as an administrator.</span></span>
    
2. <span data-ttu-id="581a8-177">Щелкните **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="581a8-177">Choose **File** > **New** > **Project**.</span></span>
    
3. <span data-ttu-id="581a8-178">В разделе **Шаблоны** разверните пункт **Visual C#** (или **Visual Basic**), затем пункт **Office/SharePoint** и выберите пункт **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="581a8-178">Under **Templates**, expand **Visual C#** (or **Visual Basic**), expand **Office/SharePoint**, and then choose **Add-ins**.</span></span>
    
4. <span data-ttu-id="581a8-179">Выберите пункт **Веб-надстройка Word**, а затем нажмите кнопку **OK**, чтобы принять параметры, используемые по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="581a8-179">Choose **Word Web Add-in**, and then choose **OK** to accept the default settings.</span></span>
       
<span data-ttu-id="581a8-180">Visual Studio создаст базовую надстройку Word, которую вы можете опубликовать в том виде, в котором она есть, не внося изменений в ее веб-проект.</span><span class="sxs-lookup"><span data-stu-id="581a8-180">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="581a8-181">Шаг 5. Публикация веб-приложения надстройки Office в Azure</span><span class="sxs-lookup"><span data-stu-id="581a8-181">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="581a8-182">Не закрывая проект вашей надстройки в Visual Studio, разверните узел решения в **обозревателе решений**, чтобы отображались оба проекта для решения.</span><span class="sxs-lookup"><span data-stu-id="581a8-182">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer** so that you see both projects for the solution.</span></span>
    
2. <span data-ttu-id="581a8-p113">Щелкните правой кнопкой мыши веб-проект и выберите пункт **Опубликовать**. Веб-проект содержит файлы веб-приложения надстройки Office, так что это именно тот проект, который вы публикуете в Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-p113">Right-click the web project and then choose **Publish**. The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>
    
3. <span data-ttu-id="581a8-185">На вкладке **Публикация** выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="581a8-185">On the **Publish** tab:</span></span>

      - <span data-ttu-id="581a8-186">Выберите пункт **Служба приложений Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="581a8-186">Choose **Microsoft Azure App Service**.</span></span>
      
      - <span data-ttu-id="581a8-187">Щелкните **Выбрать существующую**.</span><span class="sxs-lookup"><span data-stu-id="581a8-187">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="581a8-188">Щелкните **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="581a8-188">Choose **Publish**.</span></span> 

6. <span data-ttu-id="581a8-189">В диалоговом окне **Служба приложений** найдите и выберите веб-приложение, которое вы создали на [шаге 3](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="581a8-189">In the **App Service** dialog box, find and choose the web app that you created in [Step 3: Create a web app in Azure](../publish/host-an-office-add-in-on-microsoft-azure.md#step-3-create-a-web-app-in-azure) and then choose **OK**.</span></span> 

    <span data-ttu-id="581a8-p114">Visual Studio опубликует веб-проект надстройки Office в вашем веб-приложении Azure. Когда Visual Studio завершит публикацию веб-проекта, откроется браузер, в котором отобразится веб-страница с текстом "Приложение службы приложений создано". Это текущая страница, используемая по умолчанию, для веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="581a8-p114">Visual Studio publishes the web project for your Office Add-in to your Azure web app. When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created." This is the current default page for the web app.</span></span>

7. <span data-ttu-id="581a8-193">Чтобы просмотреть веб-страницу надстройки, измените URL-адрес так, чтобы он использовал HTTPS и задавал путь к HTML-странице надстройки (например: https://YourDomain.azurewebsites.net/Home.html).</span><span class="sxs-lookup"><span data-stu-id="581a8-193">To see the webpage for your add-in, change the URL so that it uses HTTPS and specifies the path of your add-in's HTML page (for example: https://YourDomain.azurewebsites.net/Home.html).</span></span> <span data-ttu-id="581a8-194">Это подтверждает, что веб-приложение надстройки теперь размещено в Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-194">This confirms that your add-in's website is now hosted on Azure.</span></span> <span data-ttu-id="581a8-195">Скопируйте корневой URL-адрес (например: https://YourDomain.azurewebsites.net). Он потребуется при редактировании манифеста на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="581a8-195">Copy this URL because you'll need it when you edit the add-in manifest file later in this topic.</span></span>
    
## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="581a8-196">Действие 6. Редактирование и развертывание XML-файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="581a8-196">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="581a8-197">В Visual Studio (с примером надстройки Office, открытом в **обозревателе решений**) разверните решение так, чтобы отображались оба проекта.</span><span class="sxs-lookup"><span data-stu-id="581a8-197">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>
    
2. <span data-ttu-id="581a8-p116">Разверните проект надстройки Office (например, WordWebAddIn), щелкните правой кнопкой мыши папку манифеста, а затем нажмите кнопку **Открыть**. Откроется XML-файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="581a8-p116">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**. The add-in XML manifest file opens.</span></span>
    
3. <span data-ttu-id="581a8-200">В XML-файле манифеста найдите и замените все экземпляры ~remoteAppUrl URL-адресом корня веб-приложения надстройки в Azure.</span><span class="sxs-lookup"><span data-stu-id="581a8-200">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure. This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> <span data-ttu-id="581a8-201">Это URL-адрес, который вы скопировали ранее после публикации веб-приложения надстройки в Azure (например, https://YourDomain.azurewebsites.net)).</span><span class="sxs-lookup"><span data-stu-id="581a8-201">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 
    
4. <span data-ttu-id="581a8-p118">Щелкните **Файл** и выберите пункт **Сохранить все**. Закройте XML-файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="581a8-p118">Choose **File** and then choose **Save All**. Close the add-in XML manifest file.</span></span>
    
5. <span data-ttu-id="581a8-204">Вернитесь в **обозреватель решений**, щелкните правой кнопкой мыши папку манифеста и выберите пункт **Открыть папку в проводнике**.</span><span class="sxs-lookup"><span data-stu-id="581a8-204">Back in **Solution Explorer**, right-click the manifest folder and choose **Open Folder In File Explorer**.</span></span>
    
6. <span data-ttu-id="581a8-205">Скопируйте XML-файл манифеста надстройки (например, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="581a8-205">Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span> 
    
7. <span data-ttu-id="581a8-206">Откройте сетевой файловый ресурс, который вы создали в [действии 1 "Создание общей папки"](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) и вставьте файл манифеста в папку.</span><span class="sxs-lookup"><span data-stu-id="581a8-206">Browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="581a8-207">Шаг 7. Вставка и запуск надстройки в клиентском приложении Office</span><span class="sxs-lookup"><span data-stu-id="581a8-207">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="581a8-208">Запустите Word и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="581a8-208">Start Word 2016 and create a document.</span></span>
    
2. <span data-ttu-id="581a8-209">На ленте щелкните **Вставка** > **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="581a8-209">On the ribbon, choose **Insert** > **My Add-ins**.</span></span> 
    
3. <span data-ttu-id="581a8-p119">В диалоговом окне **Надстройки Office** выберите **ОБЩАЯ ПАПКА**. Word выполнит сканирование папки, которую вы указали в качестве надежного каталога надстроек (в [действии 2 "Добавление файлового ресурса в надежный каталог надстроек"](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) и отобразит надстройки в диалоговом окне. Должен отобразиться значок для вашего примера надстройки.</span><span class="sxs-lookup"><span data-stu-id="581a8-p119">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**. Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box. You should see an icon for your sample add-in.</span></span>
    
4. <span data-ttu-id="581a8-p120">Щелкните значок своей надстройки и нажмите кнопку **Добавить**. На ленту будет добавлена кнопка **Показать область задач** для вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="581a8-p120">Choose the icon for your add-in and then choose **Add**. A **Show Taskpane** button for your add-in is added to the ribbon.</span></span> 

5. <span data-ttu-id="581a8-p121">На ленте вкладки **Главная** нажмите кнопку **Показать область задач**. Надстройка откроется в области задач справа от текущего документа.</span><span class="sxs-lookup"><span data-stu-id="581a8-p121">On the ribbon of the **Home** tab, choose the **Show Taskpane** button. The add-in opens in a task pane to the right of the current document.</span></span>
    
6. <span data-ttu-id="581a8-p122">Убедитесь, что надстройка работает, выбрав любой текст в документе и нажав кнопку **Highlight!** (Выделить!) в области задач.</span><span class="sxs-lookup"><span data-stu-id="581a8-p122">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!** button in the task pane.</span></span> 

## <a name="see-also"></a><span data-ttu-id="581a8-219">См. также</span><span class="sxs-lookup"><span data-stu-id="581a8-219">See also</span></span>

- [<span data-ttu-id="581a8-220">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="581a8-220">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="581a8-221">Упаковка надстройки с помощью Visual Studio для публикации</span><span class="sxs-lookup"><span data-stu-id="581a8-221">Package your add-in using Visual Studio to prepare for publishing</span></span>](../publish/package-your-add-in-using-visual-studio.md)
    
