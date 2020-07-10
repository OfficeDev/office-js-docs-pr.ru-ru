---
title: Размещение надстройки Office в Microsoft Azure | Документация Майкрософт
description: Сведения о развертывании веб-приложения надстройки в Azure и загрузке неопубликованной надстройки для тестирования в клиентском приложении Office.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: a30f1a8219501a68e6f46f013ef46640a59fe4e9
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094234"
---
# <a name="host-an-office-add-in-on-microsoft-azure"></a><span data-ttu-id="edeea-103">Размещение надстройки Office в Microsoft Azure</span><span class="sxs-lookup"><span data-stu-id="edeea-103">Host an Office Add-in on Microsoft Azure</span></span>

<span data-ttu-id="edeea-104">The simplest Office Add-in is made up of an XML manifest file and an HTML page.</span><span class="sxs-lookup"><span data-stu-id="edeea-104">The simplest Office Add-in is made up of an XML manifest file and an HTML page.</span></span> <span data-ttu-id="edeea-105">The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page.</span><span class="sxs-lookup"><span data-stu-id="edeea-105">The XML manifest file describes the add-in's characteristics, such as its name, what Office desktop applications it can run in, and the URL for the add-in's HTML page.</span></span> <span data-ttu-id="edeea-106">The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application.</span><span class="sxs-lookup"><span data-stu-id="edeea-106">The HTML page is contained in a web app that users interact with when they install and run your add-in within an Office client application.</span></span> <span data-ttu-id="edeea-107">You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span><span class="sxs-lookup"><span data-stu-id="edeea-107">You can host the web app of an Office Add-in on any web hosting platform, including Azure.</span></span>

<span data-ttu-id="edeea-108">В этой статье рассказывается, как развернуть веб-приложение надстройки в Azure и [загрузить неопубликованную надстройку](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) для тестирования в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="edeea-108">This article describes how to deploy an add-in web app to Azure and [sideload the add-in](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) for testing in an Office client application.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="edeea-109">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="edeea-109">Prerequisites</span></span> 

1. <span data-ttu-id="edeea-110">Установите [Visual Studio 2019](https://www.visualstudio.com/downloads) и не забудьте включить рабочую нагрузку **Разработка для Azure**.</span><span class="sxs-lookup"><span data-stu-id="edeea-110">Install [Visual Studio 2019](https://www.visualstudio.com/downloads) and choose to include the **Azure development** workload.</span></span>

    > [!NOTE]
    > <span data-ttu-id="edeea-111">Если Visual Studio 2019 уже установлен, убедитесь, что рабочая нагрузка **Разработка для Azure** установлена, [используя установщик Visual Studio](/visualstudio/install/modify-visual-studio).</span><span class="sxs-lookup"><span data-stu-id="edeea-111">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Azure development** workload is installed.</span></span> 

2. <span data-ttu-id="edeea-112">Установите Office.</span><span class="sxs-lookup"><span data-stu-id="edeea-112">Install Office.</span></span>

    > [!NOTE]
    > <span data-ttu-id="edeea-113">Если у вас еще нет Office, можете [оформить бесплатную пробную подписку на 1 месяц](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span><span class="sxs-lookup"><span data-stu-id="edeea-113">If you don't already have Office, you can [register for a free 1-month trial](https://products.office.com/try?legRedir=true&WT.intid1=ODC_ENUS_FX101785584_XT104056786&CorrelationId=64c762de-7a97-4dd1-bb96-e231d7485735).</span></span>

3. <span data-ttu-id="edeea-114">Подпишитесь на Azure.</span><span class="sxs-lookup"><span data-stu-id="edeea-114">Obtain an Azure subscription.</span></span>

    > [!NOTE]
    > <span data-ttu-id="edeea-115">Если у вас еще нет подписки на Azure, вы можете [получить ее в рамках своей подписки на Visual Studio](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) или [зарегистрировать бесплатную учетную запись](https://azure.microsoft.com/pricing/free-trial).</span><span class="sxs-lookup"><span data-stu-id="edeea-115">If don't already have an Azure subscription, you can [get one as part of your Visual Studio subscription](https://azure.microsoft.com/pricing/member-offers/visual-studio-subscriptions/) or [register for a free trial](https://azure.microsoft.com/pricing/free-trial).</span></span> 

## <a name="step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file"></a><span data-ttu-id="edeea-116">Шаг 1. Создание общей папки для размещения XML-файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="edeea-116">Step 1: Create a shared folder to host your add-in XML manifest file</span></span>

1. <span data-ttu-id="edeea-117">Откройте проводник на своем компьютере разработчика.</span><span class="sxs-lookup"><span data-stu-id="edeea-117">Open File Explorer on your development computer.</span></span>

2. <span data-ttu-id="edeea-118">Щелкните диск C: правой кнопкой мыши и выберите пункты **Создать** > **Папку**.</span><span class="sxs-lookup"><span data-stu-id="edeea-118">Right-click the C:\ drive and then choose **New** > **Folder**.</span></span>

3. <span data-ttu-id="edeea-119">Назовите новую папку AddinManifests.</span><span class="sxs-lookup"><span data-stu-id="edeea-119">Name the new folder AddinManifests.</span></span>

4. <span data-ttu-id="edeea-120">Щелкните папку AddinManifests правой кнопкой мыши и выберите пункты **Общий доступ** > **Конкретные пользователи...**.</span><span class="sxs-lookup"><span data-stu-id="edeea-120">Right-click the AddinManifests folder and then choose **Share with** > **Specific people**.</span></span>

5. <span data-ttu-id="edeea-121">В окне **Общий доступ к файлам** щелкните стрелку раскрывающегося списка и выберите **Все** > **Добавить** > **Общий доступ**.</span><span class="sxs-lookup"><span data-stu-id="edeea-121">In **File Sharing**, choose the drop-down arrow and then choose **Everyone** > **Add** > **Share**.</span></span>

> [!NOTE]
> <span data-ttu-id="edeea-122">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file.</span><span class="sxs-lookup"><span data-stu-id="edeea-122">In this walkthrough, you're using a local file share as a trusted catalog where you'll store the add-in XML manifest file.</span></span> <span data-ttu-id="edeea-123">In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span><span class="sxs-lookup"><span data-stu-id="edeea-123">In a real-world scenario, you might instead choose to [deploy the XML manifest file to a SharePoint catalog](../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md) or [publish the add-in to AppSource](/office/dev/store/submit-to-appsource-via-partner-center).</span></span>

## <a name="step-2-add-the-file-share-to-the-trusted-add-ins-catalog"></a><span data-ttu-id="edeea-124">Шаг 2. Добавление общей папки в доверенный каталог надстроек</span><span class="sxs-lookup"><span data-stu-id="edeea-124">Step 2: Add the file share to the Trusted Add-ins catalog</span></span>

1. <span data-ttu-id="edeea-125">Запустите Word и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="edeea-125">Start Word and create a document.</span></span>

    > [!NOTE]
    > <span data-ttu-id="edeea-126">В этом примере используется Word, но вы можете использовать любое приложение Office, поддерживающее надстройки Office, например Excel, Outlook, PowerPoint или Project.</span><span class="sxs-lookup"><span data-stu-id="edeea-126">Although this example uses Word, you can use any Office application that supports Office Add-ins such as Excel, Outlook, PowerPoint, or Project.</span></span>

2. <span data-ttu-id="edeea-127">Щелкните **Файл** > **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="edeea-127">Choose **File** > **Options**.</span></span>

3. <span data-ttu-id="edeea-128">В диалоговом окне **Параметры Word** щелкните **Центр управления безопасностью**, а затем — **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="edeea-128">In the **Word Options** dialog box, choose **Trust Center** and then choose **Trust Center Settings**.</span></span>

4. <span data-ttu-id="edeea-129">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**.</span><span class="sxs-lookup"><span data-stu-id="edeea-129">In the **Trust Center** dialog box, choose **Trusted Add-in Catalogs**.</span></span> <span data-ttu-id="edeea-130">Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span><span class="sxs-lookup"><span data-stu-id="edeea-130">Enter the universal naming convention (UNC) path for the file share you created earlier as the **Catalog URL** (for example, \\\YourMachineName\AddinManifests), and then choose **Add catalog**.</span></span> 

5. <span data-ttu-id="edeea-131">Установите флажок **Показывать в меню**.</span><span class="sxs-lookup"><span data-stu-id="edeea-131">Select the check box for **Show in Menu**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="edeea-132">Когда XML-файл манифеста надстройки хранится в доверенном каталоге веб-надстроек, надстройка отображается в разделе **Общая папка** в диалоговом окне **Надстройки Office** (**Вставка** > **Мои надстройки**).</span><span class="sxs-lookup"><span data-stu-id="edeea-132">When you store an add-in XML manifest file on a share that is specified as a trusted web add-in catalog, the add-in appears under **Shared Folder** in the **Office Add-ins** dialog box when the user navigates to the **Insert** tab in the ribbon and chooses **My Add-ins**.</span></span>

6. <span data-ttu-id="edeea-133">Закройте Word.</span><span class="sxs-lookup"><span data-stu-id="edeea-133">Close Word.</span></span>

## <a name="step-3-create-a-web-app-in-azure-using-the-azure-portal"></a><span data-ttu-id="edeea-134">Шаг 3. Создание веб-приложения в Azure с помощью портала Azure</span><span class="sxs-lookup"><span data-stu-id="edeea-134">Step 3: Create a web app in Azure using the Azure portal</span></span>

<span data-ttu-id="edeea-135">Чтобы создать веб-приложение с помощью портала Azure, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="edeea-135">To create the web app using the Azure portal, complete the following steps.</span></span>

1. <span data-ttu-id="edeea-136">Войдите в систему на [портале Azure](https://portal.azure.com/), используя свои учетные данные Azure.</span><span class="sxs-lookup"><span data-stu-id="edeea-136">Log on to the [Azure portal](https://portal.azure.com/) using your Azure credentials.</span></span>

2. <span data-ttu-id="edeea-137">В разделе**Службы Azure** выберите **Веб-приложения**.</span><span class="sxs-lookup"><span data-stu-id="edeea-137">Under **Azure Services** select **Web Apps**.</span></span>

3. <span data-ttu-id="edeea-138">На странице **Служба приложений** выберите **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="edeea-138">On the **App Service** page, select **Add**.</span></span> <span data-ttu-id="edeea-139">Чтобы добавить эти сведения, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="edeea-139">Provide this information:</span></span>

      - <span data-ttu-id="edeea-140">Выберите **подписку**, которую необходимо использовать для создания сайта.</span><span class="sxs-lookup"><span data-stu-id="edeea-140">Choose the **Subscription** to use for creating this site.</span></span>
      
      - <span data-ttu-id="edeea-141">Choose the **Resource Group** for your site.</span><span class="sxs-lookup"><span data-stu-id="edeea-141">Choose the **Resource Group** for your site.</span></span> <span data-ttu-id="edeea-142">If you create a new group, you also need to name it.</span><span class="sxs-lookup"><span data-stu-id="edeea-142">If you create a new group, you also need to name it.</span></span>
      
      - <span data-ttu-id="edeea-143">Введите уникальное **имя приложения** для своего сайта.</span><span class="sxs-lookup"><span data-stu-id="edeea-143">Enter a unique **App name** for your site.</span></span> <span data-ttu-id="edeea-144">Azure проверит уникальность имени сайта в домене azureweb apps.net.</span><span class="sxs-lookup"><span data-stu-id="edeea-144">Azure verifies that the site name is unique across the azureweb apps.net domain.</span></span>

      - <span data-ttu-id="edeea-145">Укажите, следует ли выполнить публикацию с помощью кода или контейнера Docker.</span><span class="sxs-lookup"><span data-stu-id="edeea-145">Choose whether to publish using code or a docker container.</span></span>

      - <span data-ttu-id="edeea-146">Укажите **Стек среды выполнения**.</span><span class="sxs-lookup"><span data-stu-id="edeea-146">Specify a **Runtime stack**.</span></span>

      - <span data-ttu-id="edeea-147">Выберите **операционную систему** для своего сайта.</span><span class="sxs-lookup"><span data-stu-id="edeea-147">Choose the **OS** for your site.</span></span>

      - <span data-ttu-id="edeea-148">Выберите **Регион**.</span><span class="sxs-lookup"><span data-stu-id="edeea-148">Choose a **Region**.</span></span>

      - <span data-ttu-id="edeea-149">Выберите **план службы приложений**, который необходимо использовать для создания этого сайта.</span><span class="sxs-lookup"><span data-stu-id="edeea-149">Choose the **App Service plan** to use for creating this site.</span></span>

      - <span data-ttu-id="edeea-150">Нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="edeea-150">Choose **Create**.</span></span>

4. <span data-ttu-id="edeea-151">На следующей странице вы узнаете о том, как выполняется развертывание и когда оно завершится.</span><span class="sxs-lookup"><span data-stu-id="edeea-151">The next page will let you know that your deployment is underway and when it completes.</span></span> <span data-ttu-id="edeea-152">После завершения развертывания выберите пункт **Перейти к ресурсу**.</span><span class="sxs-lookup"><span data-stu-id="edeea-152">When it is completed, select **Go to resource**.</span></span>  

5. <span data-ttu-id="edeea-153">В разделе **Обзор** выберите URL-адрес, который отображается в пункте **URL**.</span><span class="sxs-lookup"><span data-stu-id="edeea-153">In the **Overview** section, choose the URL that is displayed under **URL**.</span></span> <span data-ttu-id="edeea-154">Откроется браузер, и в нем отобразится веб-страница с сообщением "Ваша служба приложений готова к работе".</span><span class="sxs-lookup"><span data-stu-id="edeea-154">Your browser opens and displays a webpage with the message "Your App Service app is up and running."</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="edeea-155">Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="edeea-155">Azure websites automatically provide an HTTPS endpoint.</span></span>

## <a name="step-4-create-an-office-add-in-in-visual-studio"></a><span data-ttu-id="edeea-156">Шаг 4. Создание надстройки Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="edeea-156">Step 4: Create an Office Add-in in Visual Studio</span></span>

1. <span data-ttu-id="edeea-157">Запустите Visual Studio от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="edeea-157">Start Visual Studio as an administrator.</span></span>

2. <span data-ttu-id="edeea-158">Выберите **Создание нового проекта**.</span><span class="sxs-lookup"><span data-stu-id="edeea-158">Choose **Create a new project**.</span></span>

3. <span data-ttu-id="edeea-159">Используя поле поиска, введите **надстройка**.</span><span class="sxs-lookup"><span data-stu-id="edeea-159">Using the search box, enter **add-in**.</span></span>

4. <span data-ttu-id="edeea-160">Выберите пункт **Веб-надстройка Word** в качестве типа проекта, а затем нажмите кнопку **Далее**, чтобы принять параметры, используемые по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="edeea-160">Choose **Word Web Add-in** as the project type, and then choose **Next** to accept the default settings.</span></span>

<span data-ttu-id="edeea-161">Visual Studio создаст базовую надстройку Word, которую вы можете опубликовать в том виде, в котором она есть, не внося изменений в ее веб-проект.</span><span class="sxs-lookup"><span data-stu-id="edeea-161">Visual Studio creates a basic Word add-in that you'll be able to publish as-is, without making any changes to its web project.</span></span> <span data-ttu-id="edeea-162">Чтобы создать надстройку для другого типа узла Office, например Excel, повторите эти действия и выберите тип проекта с желаемым узлом Office.</span><span class="sxs-lookup"><span data-stu-id="edeea-162">To make an add-in for a different Office host type, such as Excel, repeat the steps and choose a project type with your desired Office host.</span></span>

## <a name="step-5-publish-your-office-add-in-web-app-to-azure"></a><span data-ttu-id="edeea-163">Действие 5. Публикация веб-приложения надстройки Office в Azure</span><span class="sxs-lookup"><span data-stu-id="edeea-163">Step 5: Publish your Office Add-in web app to Azure</span></span>

1. <span data-ttu-id="edeea-164">Не закрывая проект вашей надстройки в Visual Studio, разверните узел решения в **Обозревателе решений**, затем выберите **Служба приложений**.</span><span class="sxs-lookup"><span data-stu-id="edeea-164">With your add-in project open in Visual Studio, expand the solution node in **Solution Explorer**, then select **App Service**.</span></span>

2. <span data-ttu-id="edeea-165">Right-click the web project and then choose **Publish**.</span><span class="sxs-lookup"><span data-stu-id="edeea-165">Right-click the web project and then choose **Publish**.</span></span> <span data-ttu-id="edeea-166">The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span><span class="sxs-lookup"><span data-stu-id="edeea-166">The web project contains Office Add-in web app files so this is the project that you publish to Azure.</span></span>

3. <span data-ttu-id="edeea-167">На вкладке **Публикация** выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="edeea-167">On the **Publish** tab:</span></span>

      - <span data-ttu-id="edeea-168">Выберите пункт **Служба приложений Microsoft Azure**.</span><span class="sxs-lookup"><span data-stu-id="edeea-168">Choose **Microsoft Azure App Service**.</span></span>

      - <span data-ttu-id="edeea-169">Щелкните **Выбрать существующую**.</span><span class="sxs-lookup"><span data-stu-id="edeea-169">Choose **Select Existing**.</span></span>

      - <span data-ttu-id="edeea-170">Щелкните **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="edeea-170">Choose **Publish**.</span></span>

4. <span data-ttu-id="edeea-171">Visual Studio publishes the web project for your Office Add-in to your Azure web app.</span><span class="sxs-lookup"><span data-stu-id="edeea-171">Visual Studio publishes the web project for your Office Add-in to your Azure web app.</span></span> <span data-ttu-id="edeea-172">When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created."</span><span class="sxs-lookup"><span data-stu-id="edeea-172">When Visual Studio finishes publishing the web project, your browser opens and shows a webpage with the text "Your App Service app has been created."</span></span> <span data-ttu-id="edeea-173">This is the current default page for the web app.</span><span class="sxs-lookup"><span data-stu-id="edeea-173">This is the current default page for the web app.</span></span>

5. <span data-ttu-id="edeea-174">Скопируйте URL-адрес корня (пример: https://YourDomain.azurewebsites.net); он потребуется, когда вы будете редактировать файл манифеста надстройки далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="edeea-174">Copy the root URL (for example: https://YourDomain.azurewebsites.net); you'll need it when you edit the add-in manifest file later in this article.</span></span>

## <a name="step-6-edit-and-deploy-the-add-in-xml-manifest-file"></a><span data-ttu-id="edeea-175">Шаг 6. Редактирование и развертывание XML-файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="edeea-175">Step 6: Edit and deploy the add-in XML manifest file</span></span>

1. <span data-ttu-id="edeea-176">В Visual Studio (с примером надстройки Office, открытом в **обозревателе решений**) разверните решение так, чтобы отображались оба проекта.</span><span class="sxs-lookup"><span data-stu-id="edeea-176">In Visual Studio with the sample Office Add-in open in **Solution Explorer**, expand the solution so that both projects show.</span></span>

2. <span data-ttu-id="edeea-177">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**.</span><span class="sxs-lookup"><span data-stu-id="edeea-177">Expand the Office Add-in project (for example WordWebAddIn), right-click the manifest folder, and then choose **Open**.</span></span> <span data-ttu-id="edeea-178">The add-in XML manifest file opens.</span><span class="sxs-lookup"><span data-stu-id="edeea-178">The add-in XML manifest file opens.</span></span>

3. <span data-ttu-id="edeea-179">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span><span class="sxs-lookup"><span data-stu-id="edeea-179">In the XML manifest file, find and replace all instances of "~remoteAppUrl" with the root URL of the add-in web app on Azure.</span></span> <span data-ttu-id="edeea-180">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span><span class="sxs-lookup"><span data-stu-id="edeea-180">This is the URL that you copied earlier after you published the add-in web app to Azure (for example: https://YourDomain.azurewebsites.net).</span></span> 

4. <span data-ttu-id="edeea-181">Щелкните **Файл** и выберите пункт **Сохранить все**.</span><span class="sxs-lookup"><span data-stu-id="edeea-181">Choose **File** and then choose **Save All**.</span></span> <span data-ttu-id="edeea-182">Затем скопируйте XML-файл манифеста надстройки (например, WordWebAddIn.xml).</span><span class="sxs-lookup"><span data-stu-id="edeea-182">Next, Copy the add-in XML manifest file (for example, WordWebAddIn.xml).</span></span>

5. <span data-ttu-id="edeea-183">С помощью программы **Проводник** откройте сетевой файловый ресурс, который вы создали в [действии 1 "Создание общей папки"](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) и вставьте файл манифеста в папку.</span><span class="sxs-lookup"><span data-stu-id="edeea-183">Using the **File Explorer** program, browse to the network file share that you created in [Step 1: Create a shared folder](../publish/host-an-office-add-in-on-microsoft-azure.md#step-1-create-a-shared-folder-to-host-your-add-in-xml-manifest-file) and paste the manifest file into the folder.</span></span>

## <a name="step-7-insert-and-run-the-add-in-in-the-office-client-application"></a><span data-ttu-id="edeea-184">Шаг 7. Вставка и запуск надстройки в клиентском приложении Office</span><span class="sxs-lookup"><span data-stu-id="edeea-184">Step 7: Insert and run the add-in in the Office client application</span></span>

1. <span data-ttu-id="edeea-185">Запустите Word и создайте документ.</span><span class="sxs-lookup"><span data-stu-id="edeea-185">Start Word and create a document.</span></span>

2. <span data-ttu-id="edeea-186">На ленте щелкните **Вставка** > **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="edeea-186">On the ribbon, choose **Insert** > **My Add-ins**.</span></span>

3. <span data-ttu-id="edeea-187">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**.</span><span class="sxs-lookup"><span data-stu-id="edeea-187">In the **Office Add-ins** dialog box, choose **SHARED FOLDER**.</span></span> <span data-ttu-id="edeea-188">Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box.</span><span class="sxs-lookup"><span data-stu-id="edeea-188">Word scans the folder that you listed as a trusted add-ins catalog (in [Step 2: Add the file share to the Trusted Add-ins catalog](../publish/host-an-office-add-in-on-microsoft-azure.md#step-2-add-the-file-share-to-the-trusted-add-ins-catalog)) and shows the add-ins in the dialog box.</span></span> <span data-ttu-id="edeea-189">You should see an icon for your sample add-in.</span><span class="sxs-lookup"><span data-stu-id="edeea-189">You should see an icon for your sample add-in.</span></span>

4. <span data-ttu-id="edeea-190">Choose the icon for your add-in and then choose **Add**.</span><span class="sxs-lookup"><span data-stu-id="edeea-190">Choose the icon for your add-in and then choose **Add**.</span></span> <span data-ttu-id="edeea-191">A **Show Taskpane** button for your add-in is added to the ribbon.</span><span class="sxs-lookup"><span data-stu-id="edeea-191">A **Show Taskpane** button for your add-in is added to the ribbon.</span></span>

5. <span data-ttu-id="edeea-192">On the ribbon of the **Home** tab, choose the **Show Taskpane** button.</span><span class="sxs-lookup"><span data-stu-id="edeea-192">On the ribbon of the **Home** tab, choose the **Show Taskpane** button.</span></span> <span data-ttu-id="edeea-193">The add-in opens in a task pane to the right of the current document.</span><span class="sxs-lookup"><span data-stu-id="edeea-193">The add-in opens in a task pane to the right of the current document.</span></span>

6. <span data-ttu-id="edeea-194">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!**</span><span class="sxs-lookup"><span data-stu-id="edeea-194">Verify that the add-in works by selecting some text in the document and choosing the **Highlight!**</span></span> <span data-ttu-id="edeea-195">button in the task pane.</span><span class="sxs-lookup"><span data-stu-id="edeea-195">button in the task pane.</span></span>

## <a name="see-also"></a><span data-ttu-id="edeea-196">См. также</span><span class="sxs-lookup"><span data-stu-id="edeea-196">See also</span></span>

- [<span data-ttu-id="edeea-197">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="edeea-197">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="edeea-198">Публикация надстройки с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="edeea-198">Publish your add-in using Visual Studio</span></span>](../publish/package-your-add-in-using-visual-studio.md)
