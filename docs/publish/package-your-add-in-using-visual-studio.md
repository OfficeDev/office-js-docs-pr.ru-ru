---
title: Упаковка надстройки с помощью Visual Studio для публикации | Документация Майкрософт
description: Способ развертывания веб-проекта и упаковки надстройки с помощью Visual Studio 2017.
ms.date: 01/25/2018
localization_priority: Priority
ms.openlocfilehash: a135e8e72703c3de60290a9eb7b2e03c63449124
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/23/2019
ms.locfileid: "29386437"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="a37d4-103">Упаковка надстройки с помощью Visual Studio для публикации</span><span class="sxs-lookup"><span data-stu-id="a37d4-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="a37d4-104">Пакет надстройки Office содержит [XML-файл манифеста](../develop/add-in-manifests.md), который будет использоваться для публикации надстройки.</span><span class="sxs-lookup"><span data-stu-id="a37d4-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="a37d4-105">Файлы веб-приложения из проекта потребуется публиковать отдельно.</span><span class="sxs-lookup"><span data-stu-id="a37d4-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="a37d4-106">В этой статье описано развертывание веб-проекта и упаковка надстройки с помощью Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="a37d4-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2017.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2017"></a><span data-ttu-id="a37d4-107">Развертывание веб-проекта с помощью Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="a37d4-107">To deploy your web project using Visual Studio 2017</span></span>

<span data-ttu-id="a37d4-108">Выполните указанные ниже действия, чтобы развернуть веб-проект с помощью Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="a37d4-108">Complete the following steps to deploy your web project using Visual Studio 2017.</span></span>

1. <span data-ttu-id="a37d4-109">В **обозревателе решений** откройте контекстное меню для проекта надстройки и выберите пункт **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="a37d4-110">Откроется страница **Опубликовать надстройку**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="a37d4-111">В раскрывающемся списке **Текущий профиль** выберите профиль или команду **Создать…**, чтобы создать профиль.</span><span class="sxs-lookup"><span data-stu-id="a37d4-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="a37d4-112">В профиле публикации указываются сервер, на котором выполняется развертывание, учетные данные для входа на сервер, развертываемые базы данных и другие параметры развертывания.</span><span class="sxs-lookup"><span data-stu-id="a37d4-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="a37d4-113">Если выбрать команду **Создать…**, откроется мастер со страницей **создания профилей публикации**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-113">If you choose  **New ...**, a wizard appears with the **Create publishing profile** page.</span></span> <span data-ttu-id="a37d4-114">С помощью этого мастера можно импортировать профиль публикации из поставщика услуг размещения веб-сайтов, например Microsoft Azure, или создать новый профиль, а затем добавить сервер, учетные данные и другие параметры, указанные на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="a37d4-114">You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="a37d4-115">Дополнительные сведения об импорте существующих профилей публикации и создании новых см. в разделе [Создание профиля публикации](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="a37d4-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="a37d4-116">На странице **Опубликовать надстройку** перейдите по ссылке **Развернуть веб-проект**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-116">On the **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="a37d4-117">Откроется диалоговое окно **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-117">The  **Publish** dialog box appears.</span></span> <span data-ttu-id="a37d4-118">Дополнительные сведения об использовании этого мастера см. в статье [Как развернуть веб-проект с использованием возможности публикации по щелчку в Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="a37d4-118">For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2017"></a><span data-ttu-id="a37d4-119">Упаковка надстройки с помощью Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="a37d4-119">To package your add-in using Visual Studio 2017</span></span>

<span data-ttu-id="a37d4-120">Выполните указанные ниже действия, чтобы упаковать надстройку с помощью Visual Studio 2017.</span><span class="sxs-lookup"><span data-stu-id="a37d4-120">Complete the following steps to package your add-in using Visual Studio 2017.</span></span>

1. <span data-ttu-id="a37d4-121">На странице **Опубликовать надстройку** нажмите кнопку **Упаковать надстройку**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-121">In the **Publish your add-in** page, choose the **Package the add-in** button.</span></span>
    
    <span data-ttu-id="a37d4-122">Появится мастер со страницей **Упаковать надстройку**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-122">A wizard appears with the **Package the add-in** page.</span></span>
    
2. <span data-ttu-id="a37d4-123">В раскрывающемся списке **Где размещается веб-сайт?** выберите или введите URL-адрес веб-сайта, на котором будут размещены файлы содержимого надстройки, а затем нажмите кнопку **Готово**.</span><span class="sxs-lookup"><span data-stu-id="a37d4-123">In the **Where is your website hosted?** box, enter the URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span>
    
    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)] <span data-ttu-id="a37d4-124">Веб-сайты Azure автоматически предоставляют конечную точку HTTPS.</span><span class="sxs-lookup"><span data-stu-id="a37d4-124">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="a37d4-125">Visual Studio создает файлы, необходимые для публикации надстройки, а затем открывает папку с выходными файлами публикации.</span><span class="sxs-lookup"><span data-stu-id="a37d4-125">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span>
    
<span data-ttu-id="a37d4-126">Если вы собираетесь отправить надстройку в AppSource, можно нажать кнопку **Выполнить проверку правильности**, чтобы определить проблемы, препятствующие приему надстройки.</span><span class="sxs-lookup"><span data-stu-id="a37d4-126">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted.</span></span> <span data-ttu-id="a37d4-127">Прежде чем отправлять надстройку в магазин, необходимо устранить все эти проблемы.</span><span class="sxs-lookup"><span data-stu-id="a37d4-127">You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="a37d4-p105">Теперь вы можете отправить XML-манифест в нужное расположение, чтобы [опубликовать надстройку](../publish/publish.md). XML-манифест можно найти в дочерней папке `OfficeAppManifests` папки `app.publish`. Например:</span><span class="sxs-lookup"><span data-stu-id="a37d4-p105">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2017\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="a37d4-131">См. также</span><span class="sxs-lookup"><span data-stu-id="a37d4-131">See also</span></span>

- [<span data-ttu-id="a37d4-132">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="a37d4-132">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="a37d4-133">Публикация решений в AppSource и в Office</span><span class="sxs-lookup"><span data-stu-id="a37d4-133">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
