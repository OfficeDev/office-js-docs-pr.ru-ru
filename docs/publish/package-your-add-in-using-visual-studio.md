---
title: Публикация надстройки с помощью Visual Studio
description: Способ развертывания веб-проекта и упаковки надстройки с помощью Visual Studio 2019.
ms.date: 12/02/2019
localization_priority: Normal
ms.openlocfilehash: 49b8b53b665b887e4f8dba20e085c3350e7711f8
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612054"
---
# <a name="publish-your-add-in-using-visual-studio"></a><span data-ttu-id="3ff0a-103">Публикация надстройки с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="3ff0a-103">Publish your add-in using Visual Studio</span></span>

<span data-ttu-id="3ff0a-104">Пакет надстройки Office содержит [XML-файл манифеста](../develop/add-in-manifests.md), который будет использоваться для публикации надстройки.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="3ff0a-105">Файлы веб-приложения из проекта потребуется публиковать отдельно.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="3ff0a-106">В этой статье описано развертывание веб-проекта и упаковка надстройки с помощью Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2019.</span></span>

> [!NOTE]
> <span data-ttu-id="3ff0a-107">Сведения о публикации надстройки Office, созданной с помощью генератора Yeoman и разработанной в Visual Studio Code или любом другом редакторе, см. в статье [Публикация надстройки, разработанной с помощью Visual Studio Code](publish-add-in-vs-code.md).</span><span class="sxs-lookup"><span data-stu-id="3ff0a-107">For information about publishing an Office Add-in that you created using the Yeoman generator and developed with Visual Studio Code or any other editor, see [Publish an add-in developed with Visual Studio Code](publish-add-in-vs-code.md).</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="3ff0a-108">Развертывание веб-проекта с помощью Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="3ff0a-108">To deploy your web project using Visual Studio 2019</span></span>

<span data-ttu-id="3ff0a-109">Выполните указанные ниже действия, чтобы развернуть веб-проект с помощью Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-109">Complete the following steps to deploy your web project using Visual Studio 2019.</span></span>

1. <span data-ttu-id="3ff0a-110">На вкладке **Сборка** выберите \*\*Опубликовать [имя надстройки] \*\*.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-110">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="3ff0a-111">В диалоговом окне **Выбрать целевой объект публикации** выберите один из вариантов публикации в предпочитаемом целевом объекте.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-111">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="3ff0a-112">Для каждого целевого объекта публикации необходимо включить дополнительные сведения, чтобы начать работу, например виртуальную машину Azure или расположение папки.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-112">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="3ff0a-113">После того как вы указали место публикации и заполнили все необходимые сведения, выберите пункт **Опубликовать**</span><span class="sxs-lookup"><span data-stu-id="3ff0a-113">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="3ff0a-114">В выборе целевого объекта публикации указываются сервер, на котором выполняется развертывание, учетные данные для входа на сервер, развертываемые базы данных и другие параметры развертывания.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-114">Picking a publish target specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="3ff0a-115">Дополнительные сведения о действиях, которые необходимо выполнить для каждого целевого объекта публикации, см. в статье [Знакомство с развертыванием в Visual Studio ](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span><span class="sxs-lookup"><span data-stu-id="3ff0a-115">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="3ff0a-116">Упаковка и публикация надстройки с помощью IIS, FTP или веб-развертывания с использованием Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="3ff0a-116">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="3ff0a-117">Выполните указанные ниже действия, чтобы упаковать надстройку с помощью Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-117">Complete the following steps to package your add-in using Visual Studio 2019.</span></span>

1. <span data-ttu-id="3ff0a-118">На вкладке **Сборка** выберите \*\*Опубликовать [имя надстройки] \*\*.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-118">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="3ff0a-119">В окне **Выбрать целевой объект публикации** выберите **IIS, FTP и т. д.**, затем выберите **Настроить**.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-119">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="3ff0a-120">Далее нажмите **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-120">Next, select **Publish**.</span></span>
3. <span data-ttu-id="3ff0a-121">Откроется мастер, который поможет вам выполнить все необходимые действия.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-121">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="3ff0a-122">Убедитесь в том, что метод публикации является вашим предпочтительным методом, таким как веб-развертывание.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-122">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="3ff0a-123">В поле **Целевой URL-адрес** введите URL-адрес веб-сайта, на котором будут размещены файлы содержимого надстройки, а затем нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-123">In the **Destination URL** box, enter the URL of the website that will host the content files of your add-in, and then select **Next**.</span></span> <span data-ttu-id="3ff0a-124">Если вы собираетесь отправить надстройку в AppSource, можно нажать кнопку **Проверить подключение**, чтобы определить проблемы, препятствующие приему надстройки.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-124">If you plan to submit your add-in to AppSource, you can choose the **Validate Connection** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="3ff0a-125">Прежде чем отправлять надстройку в магазин, необходимо устранить все эти проблемы.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-125">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="3ff0a-126">Подтвердите любые желаемые настройки, включая **Варианты публикации файла** и выберите **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-126">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="3ff0a-127">Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="3ff0a-127">Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="3ff0a-p106">Теперь вы можете отправить XML-манифест в нужное расположение, чтобы [опубликовать надстройку](../publish/publish.md). XML-манифест можно найти в дочерней папке `OfficeAppManifests` папки `app.publish`. Например:</span><span class="sxs-lookup"><span data-stu-id="3ff0a-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="3ff0a-131">См. также</span><span class="sxs-lookup"><span data-stu-id="3ff0a-131">See also</span></span>

- [<span data-ttu-id="3ff0a-132">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="3ff0a-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="3ff0a-133">Публикация решений в AppSource и в Office</span><span class="sxs-lookup"><span data-stu-id="3ff0a-133">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
