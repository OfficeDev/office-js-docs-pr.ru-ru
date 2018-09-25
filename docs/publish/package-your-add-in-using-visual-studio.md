---
title: Упаковка надстройки с помощью Visual Studio для ее подготовки к публикации | Microsoft Docs
description: Как развернуть веб-проект и выполнить упаковку надстройки с помощью Visual Studio 2015.
ms.date: 01/25/2018
ms.openlocfilehash: d74ead03b8ac5b7652c7c98851e7e082f4b31ba8
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2018
ms.locfileid: "25004919"
---
# <a name="package-your-add-in-using-visual-studio-to-prepare-for-publishing"></a><span data-ttu-id="357f0-103">Как развернуть веб-проект и выполнить упаковку надстройки с помощью Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="357f0-103">Package your add-in using Visual Studio to prepare for publishing</span></span>

<span data-ttu-id="357f0-104">Пакет надстройки Office содержит [XML-файл манифеста](../develop/add-in-manifests.md), который будет использоваться для публикации надстройки.</span><span class="sxs-lookup"><span data-stu-id="357f0-104">Your Office Add-in package contains an XML [manifest file](../develop/add-in-manifests.md) that you'll use to publish the add-in.</span></span> <span data-ttu-id="357f0-105">Файлы веб-приложения из проекта потребуется публиковать отдельно.</span><span class="sxs-lookup"><span data-stu-id="357f0-105">You'll have to publish the web application files of your project separately.</span></span> <span data-ttu-id="357f0-106">В этой статье описано развертывание веб-проекта и упаковка надстройки с помощью Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="357f0-106">This article describes how to deploy your web project and package your add-in by using Visual Studio 2015.</span></span>

## <a name="to-deploy-your-web-project-using-visual-studio-2015"></a><span data-ttu-id="357f0-107">Развертывание веб-проекта с помощью Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="357f0-107">To deploy your web project using Visual Studio 2015</span></span>

<span data-ttu-id="357f0-108">Выполните указанные ниже действия, чтобы развернуть веб-проект с помощью Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="357f0-108">Complete the following steps to deploy your web project using Visual Studio 2015.</span></span>

1. <span data-ttu-id="357f0-109">В **обозревателе решений** откройте контекстное меню для проекта надстройки и выберите пункт **Опубликовать**.</span><span class="sxs-lookup"><span data-stu-id="357f0-109">In  **Solution Explorer**, open the shortcut menu for the add-in project, and then choose  **Publish**.</span></span>
    
    <span data-ttu-id="357f0-110">Откроется страница **Опубликовать надстройку**.</span><span class="sxs-lookup"><span data-stu-id="357f0-110">The  **Publish your add-in** page appears.</span></span>
    
2. <span data-ttu-id="357f0-111">В раскрывающемся списке **Текущий профиль** выберите профиль или команду **Создать…**, чтобы создать профиль.</span><span class="sxs-lookup"><span data-stu-id="357f0-111">In the  **Current profile** drop-down list, select a profile or choose **New ...** to create a new profile.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="357f0-112">В профиле публикации указываются сервер, на котором выполняется развертывание, учетные данные для входа на сервер, развертываемые базы данных и другие параметры развертывания.</span><span class="sxs-lookup"><span data-stu-id="357f0-112">A publish profile specifies the server you are deploying to, the credentials needed to log on to the server, the databases to deploy, and other deployment options.</span></span>

    <span data-ttu-id="357f0-113">Если выбрать опцию  **Новый ...**, то откроется мастер создания профилей публикации.</span><span class="sxs-lookup"><span data-stu-id="357f0-113">If you choose  New ..., the  Create publishing profile wizard appears.</span></span> <span data-ttu-id="357f0-114">С помощью этого мастера можно импортировать профиль публикации из поставщика услуг размещения веб-сайтов, например Microsoft Azure, или создать новый профиль, а затем добавить сервер, учетные данные и другие параметры, указанные на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="357f0-114">You can use this wizard to import a publishing profile from a web site hosting provider such as Microsoft Azure or create a new profile and add your server, credentials, and other settings in the next procedure.</span></span>
    
    <span data-ttu-id="357f0-115">Дополнительные сведения об импорте существующих профилей публикации и создании новых см. в разделе [Создание профиля публикации](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span><span class="sxs-lookup"><span data-stu-id="357f0-115">For more information about importing publishing profiles or creating new publishing profiles, see [Creating a Publish Profile](https://msdn.microsoft.com/library/dd465337.aspx#creating_a_profile).</span></span>
    
3. <span data-ttu-id="357f0-116">На странице  **Публикация надстройки** перейдите по ссылке **Выполнить развертывание веб-проекта**.</span><span class="sxs-lookup"><span data-stu-id="357f0-116">In the  **Publish your add-in** page, choose the **Deploy your web project** link.</span></span>
    
    <span data-ttu-id="357f0-p103">Появится диалоговое окно  **Опубликовать веб-проект**. Более подробную информацию об использовании этого мастера см. в статье [Развертывание веб-проекта с помощью публикации одним щелчком в Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span><span class="sxs-lookup"><span data-stu-id="357f0-p103">The  **Publish Web** dialog box appears. For more information about using this wizard, see [How to: Deploy a Web Project using On-Click Publishing in Visual Studio](https://msdn.microsoft.com/library/dd465337.aspx).</span></span>
    

## <a name="to-package-your-add-in-using-visual-studio-2015"></a><span data-ttu-id="357f0-119">Упаковка надстройки с помощью Visual Studio 2015</span><span class="sxs-lookup"><span data-stu-id="357f0-119">To package your add-in using Visual Studio 2015</span></span>

<span data-ttu-id="357f0-120">Выполните указанные ниже действия, чтобы упаковать надстройку с помощью Visual Studio 2015.</span><span class="sxs-lookup"><span data-stu-id="357f0-120">Complete the following steps to package your add-in using Visual Studio 2015.</span></span>

1. <span data-ttu-id="357f0-121">На странице **Публикация надстройки** перейдите по ссылке **Упаковка надстройки**.</span><span class="sxs-lookup"><span data-stu-id="357f0-121">In the **Publish your add-in** page, choose the **Package the add-in** link.</span></span>
    
    <span data-ttu-id="357f0-122">При этом откроется мастер публикации надстроек Office и SharePoint.</span><span class="sxs-lookup"><span data-stu-id="357f0-122">The Publish Office and SharePoint Add-ins wizard appears.</span></span>
    
2. <span data-ttu-id="357f0-123">В раскрывающемся списке **Где размещается веб-сайт?** выберите или введите URL-адрес HTTPS веб-сайта, на котором будут размещены файлы содержимого надстройки, а затем нажмите кнопку **Готово**.</span><span class="sxs-lookup"><span data-stu-id="357f0-123">In the **Where is your website hosted?** dropdown list, select or enter the HTTPS URL of the website that will host the content files of your add-in, and then choose **Finish**.</span></span> 
    
    <span data-ttu-id="357f0-p104">Для успешного завершения работы этого мастера необходимо указать URL-адрес с префиксом HTTPS. Если вы хотите использовать конечную точку HTTP для веб-сайта, можно открыть XML-файл манифеста в текстовом редакторе после создания пакета и заменить префикс HTTPS веб-сайта на префикс HTTP.</span><span class="sxs-lookup"><span data-stu-id="357f0-p104">You must specify a URL that begins with the HTTPS prefix to complete this wizard. If you want to use an HTTP endpoint for your website, you can open the XML manifest file in a text editor after the package has been created and replace the HTTPS prefix of your website with an HTTP prefix.</span></span> 

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]<span data-ttu-id="357f0-126"> Сайтами Azure автоматически предоставляется конечная точка HTTPS.</span><span class="sxs-lookup"><span data-stu-id="357f0-126">Azure websites automatically provide an HTTPS endpoint.</span></span>

    <span data-ttu-id="357f0-127">Visual Studio создает файлы, необходимые для публикации надстройки, а затем открывает папку с выходными файлами публикации.</span><span class="sxs-lookup"><span data-stu-id="357f0-127">Visual Studio generates the files that you need to publish your add-in and then opens the publish output folder.</span></span> 
    
<span data-ttu-id="357f0-p105">Если вы собираетесь отправить надстройку в AppSource, можете выбрать ссылку **Выполните проверку правильности**, чтобы определить проблемы, препятствующие приему надстройки. Перед отправкой надстройки в магазин необходимо решить все проблемы.</span><span class="sxs-lookup"><span data-stu-id="357f0-p105">If you plan to submit your add-in to AppSource, you can choose the **Perform a validation check** link to identify any issues that will prevent your add-in from being accepted. You should address all issues before you submit your add-in to the store.</span></span>

<span data-ttu-id="357f0-p106">Теперь вы можете отправить XML-манифест в нужное расположение, чтобы [опубликовать надстройку](../publish/publish.md). XML-манифест можно найти в дочерней папке `OfficeAppManifests` папки `app.publish`. Например:</span><span class="sxs-lookup"><span data-stu-id="357f0-p106">You can now upload your XML manifest to the appropriate location to [publish your add-in](../publish/publish.md). You can find the XML manifest in `OfficeAppManifests` in the `app.publish` folder. For example:</span></span>

 `%UserProfile%\Documents\Visual Studio 2015\Projects\MyApp\bin\Debug\app.publish\OfficeAppManifests`


## <a name="see-also"></a><span data-ttu-id="357f0-133">См. также</span><span class="sxs-lookup"><span data-stu-id="357f0-133">See also</span></span>

- [<span data-ttu-id="357f0-134">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="357f0-134">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="357f0-135">Публикация решений в AppSource и в Office</span><span class="sxs-lookup"><span data-stu-id="357f0-135">Make your solutions available in AppSource and within Office</span></span>](https://docs.microsoft.com/office/dev/store/submit-to-the-office-store)
    
