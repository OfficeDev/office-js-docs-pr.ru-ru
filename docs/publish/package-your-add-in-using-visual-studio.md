---
title: Упаковка надстройки с помощью Visual Studio для публикации
description: Способ развертывания веб-проекта и упаковки надстройки с помощью Visual Studio 2019.
ms.date: 10/14/2019
localization_priority: Priority
ms.openlocfilehash: 784741cffa0e3015caaa9c70fbb56f4b70df9462
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626966"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="e06bf-103">Упаковка надстройки с помощью Visual Studio для публикации</span><span class="sxs-lookup"><span data-stu-id="e06bf-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="e06bf-104">Пакет надстройки Office содержит [XML-файл манифеста](../develop/add-in-manifests.md), который будет использоваться для публикации надстройки.</span><span class="sxs-lookup"><span data-stu-id="e06bf-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="e06bf-105">Файлы веб-приложения из проекта потребуется публиковать отдельно.</span><span class="sxs-lookup"><span data-stu-id="e06bf-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="e06bf-106">В этой статье описано развертывание веб-проекта и упаковка надстройки с помощью Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="e06bf-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2019"></a><span data-ttu-id="e06bf-107">Развертывание веб-проекта с помощью Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="e06bf-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="e06bf-108">Выполните указанные ниже действия, чтобы развернуть веб-проект с помощью Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="e06bf-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="e06bf-109">На вкладке **Сборка** выберите \*\*Опубликовать [имя надстройки] \*\*.</span><span class="sxs-lookup"><span data-stu-id="e06bf-109">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>

2. <span data-ttu-id="e06bf-110">В диалоговом окне **Выбрать целевой объект публикации** выберите один из вариантов публикации в предпочитаемом целевом объекте.</span><span class="sxs-lookup"><span data-stu-id="e06bf-110">In the **Pick a publish target** window, choose one of the options to publish to your preferred target.</span></span> <span data-ttu-id="e06bf-111">Для каждого целевого объекта публикации необходимо включить дополнительные сведения, чтобы начать работу, например виртуальную машину Azure или расположение папки.</span><span class="sxs-lookup"><span data-stu-id="e06bf-111">Each publish target requires you to include more information to get started, such as an Azure Virtual Machine or folder location.</span></span> <span data-ttu-id="e06bf-112">После того как вы указали место публикации и заполнили все необходимые сведения, выберите пункт **Опубликовать**</span><span class="sxs-lookup"><span data-stu-id="e06bf-112">Once you have specified a publish location and filled in all of the information required, select **Publish**</span></span>

    > [!NOTE]
    > <span data-ttu-id="e06bf-113">В выборе целевого объекта публикации указываются сервер, на котором выполняется развертывание, учетные данные для входа на сервер, развертываемые базы данных и другие параметры развертывания.</span><span class="sxs-lookup"><span data-stu-id="e06bf-113">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

3. <span data-ttu-id="e06bf-114">Дополнительные сведения о действиях, которые необходимо выполнить для каждого целевого объекта публикации, см. в статье [Знакомство с развертыванием в Visual Studio ](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span><span class="sxs-lookup"><span data-stu-id="e06bf-114">For more information about deployment steps for each publish target option, see [First look at deployment in Visual Studio](/visualstudio/deployment/deploying-applications-services-and-components?view=vs-2019).</span></span>

## <a name="to-package-and-publish-your-add-in-using-iis-ftp-or-web-deploy-using-visual-studio-2019"></a><span data-ttu-id="e06bf-115">Упаковка и публикация надстройки с помощью IIS, FTP или веб-развертывания с использованием Visual Studio 2019</span><span class="sxs-lookup"><span data-stu-id="e06bf-115">To package and publish your add-in using IIS, FTP, or Web Deploy using Visual Studio 2019</span></span>

<span data-ttu-id="e06bf-116">Выполните указанные ниже действия, чтобы упаковать надстройку с помощью Visual Studio 2019.</span><span class="sxs-lookup"><span data-stu-id="e06bf-116">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="e06bf-117">На вкладке **Сборка** выберите \*\*Опубликовать [имя надстройки] \*\*.</span><span class="sxs-lookup"><span data-stu-id="e06bf-117">From the **Build** tab, choose **Publish [Name of your add-in]**.</span></span>
2. <span data-ttu-id="e06bf-118">В окне **Выбрать целевой объект публикации** выберите **IIS, FTP и т. д.**, затем выберите **Настроить**.</span><span class="sxs-lookup"><span data-stu-id="e06bf-118">In the **Pick a publish target** window, choose **IIS, FTP, etc**, and select **Configure**.</span></span> <span data-ttu-id="e06bf-119">Далее нажмите **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="e06bf-119">Next, select **Publish**.</span></span>
3. <span data-ttu-id="e06bf-120">Откроется мастер, который поможет вам выполнить все необходимые действия.</span><span class="sxs-lookup"><span data-stu-id="e06bf-120">A wizard appears that will help guide you through the process.</span></span> <span data-ttu-id="e06bf-121">Убедитесь в том, что метод публикации является вашим предпочтительным методом, таким как веб-развертывание.</span><span class="sxs-lookup"><span data-stu-id="e06bf-121">Ensure the publish method is your preferred method, such as Web Deploy.</span></span>
4. <span data-ttu-id="e06bf-122">В поле **Целевой URL-адрес** введите URL-адрес веб-сайта, на котором будут размещены файлы содержимого надстройки, а затем нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="e06bf-122">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> <span data-ttu-id="e06bf-123">Если вы собираетесь отправить надстройку в AppSource, можно нажать кнопку **Проверить подключение**, чтобы определить проблемы, препятствующие приему надстройки.</span><span class="sxs-lookup"><span data-stu-id="e06bf-123">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** button to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="e06bf-124">Прежде чем отправлять надстройку в магазин, необходимо устранить все эти проблемы.</span><span class="sxs-lookup"><span data-stu-id="e06bf-124">You should address all issues before you submit your add-in to the store.</span></span>
5. <span data-ttu-id="e06bf-125">Подтвердите любые желаемые настройки, включая **Варианты публикации файла** и выберите **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="e06bf-125">Confirm any settings desired including **File Publish Options** and select **Save**.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="e06bf-126">Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="e06bf-126">Azure websites automatically provide an HTTPS endpoint.</span></span>

<span data-ttu-id="e06bf-p106">Теперь вы можете отправить XML-манифест в нужное расположение, чтобы [опубликовать надстройку](../publish/publish.md). XML-манифест можно найти в дочерней папке `OfficeAppManifests` папки `app.publish`. Например:</span><span class="sxs-lookup"><span data-stu-id="e06bf-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2019\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`

## <a name="see-also"></a><span data-ttu-id="e06bf-130">См. также</span><span class="sxs-lookup"><span data-stu-id="e06bf-130">See also</span></span>

- [<span data-ttu-id="e06bf-131">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="e06bf-131">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="e06bf-132">Публикация решений в AppSource и в Office</span><span class="sxs-lookup"><span data-stu-id="e06bf-132">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
