---
title: Загрузка неопубликованных надстроек Office для тестирования
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: 1bca17808deaa5e7f0c65669a87abe1b38e5393f
ms.sourcegitcommit: 0d4d78e275249f0d4b6a6cf807b42b79890c3023
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2018
ms.locfileid: "21773582"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="965b0-102">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="965b0-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="965b0-103">Вы можете установить надстройку Office для тестирования в клиенте Office, работающем на Windows, опубликовав манифест в сетевом файловом ресурсе (инструкции ниже).</span><span class="sxs-lookup"><span data-stu-id="965b0-103">You can install an Office Add-in for testing in an Office client running on Windows by using a shared folder catalog to publish the manifest to a network file share.</span></span>

> [!NOTE]
> <span data-ttu-id="965b0-104">Если ваш проект надстройки был создан с помощью [**инструмента**yo office](https://github.com/OfficeDev/generator-office), есть альтернативный способ его загрузки, который может вам подойти.</span><span class="sxs-lookup"><span data-stu-id="965b0-104">If your add-in project was created with the [**yo office** tool](https://github.com/OfficeDev/generator-office), there is an alternative way of sideloading it that might work for you.</span></span> <span data-ttu-id="965b0-105">Подробнее см. в статье [Загрузка неопубликованных надстроек Office с использованием команды sideload](sideload-office-addin-using-sideload-command.md).</span><span class="sxs-lookup"><span data-stu-id="965b0-105">Sideload Office Add-ins using the sideload command</span></span>

<span data-ttu-id="965b0-106">Эта статья применяется только к тестированию надстроек Word, Excel или PowerPoint на Windows.</span><span class="sxs-lookup"><span data-stu-id="965b0-106">This article applies only to testing a Word, Excel, or PowerPoint add-ins on Windows.</span></span> <span data-ttu-id="965b0-107">Если вы хотите выполнить тестирование на другой платформе или протестировать надстройку Outlook, см. одну из следующих тем по загрузке надстройки:</span><span class="sxs-lookup"><span data-stu-id="965b0-107">If you want to test on another platform or want to test an Outlook add-in, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="965b0-108">Загрузка неопубликованных надстроек Office в Office Online для тестирования</span><span class="sxs-lookup"><span data-stu-id="965b0-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="965b0-109">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="965b0-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)
- [<span data-ttu-id="965b0-110">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="965b0-110">Sideload Outlook add-ins for testing</span></span>](../../../../outlook/add-insSideload Outlook Add-ins for testing)

<span data-ttu-id="965b0-111">В следующем видео вы ознакомитесь с процессом загрузки вашей надстройки для настольной системы Office или Office Online с помощью каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="965b0-111">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="965b0-112">Общий доступ к папке</span><span class="sxs-lookup"><span data-stu-id="965b0-112">Share a folder</span></span>

1. <span data-ttu-id="965b0-113">На том компьютере с Windows, где должна размещаться надстройка, перейдите к родительской папке или диску с папкой, которую требуется использовать в качестве каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="965b0-113">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="965b0-114">Откройте контекстное меню папки (щелкните ее правой кнопкой мыши) и выберите пункт **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="965b0-114">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="965b0-115">Откройте вкладку **Доступ**.</span><span class="sxs-lookup"><span data-stu-id="965b0-115">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="965b0-p103">На странице **Выбор людей** добавьте себя и других пользователей, которым требуется предоставить доступ к надстройке. Если все эти пользователи являются участниками группы безопасности, вы можете добавить группу. Вам потребуются разрешения на **чтение и запись** папки.</span><span class="sxs-lookup"><span data-stu-id="965b0-p103">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="965b0-119">Нажмите кнопки **Общий доступ** > **Готово** > **Закрыть**.</span><span class="sxs-lookup"><span data-stu-id="965b0-119">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="965b0-120">Укажите общую папку в качестве каталога общих папок</span><span class="sxs-lookup"><span data-stu-id="965b0-120">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="965b0-121">Откройте новый документ в Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="965b0-121">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="965b0-122">Перейдите на вкладку **Файл**, а затем выберите **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="965b0-122">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="965b0-123">Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="965b0-123">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="965b0-124">Выберите пункт **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="965b0-124">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="965b0-125">В поле **URL-адрес каталога** введите полный сетевой путь к каталогу общих папок и нажмите **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="965b0-125">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="965b0-126">Установите флажок **Показать в меню** и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="965b0-126">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="965b0-127">Закройте приложение Office, чтобы изменения вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="965b0-127">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="965b0-128">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="965b0-128">Sideload your add-in</span></span>

1. <span data-ttu-id="965b0-p104">Файл манифеста тестируемой надстройки необходимо поместить в каталог общих папок. Обратите внимание, что вы развертываете веб-приложение непосредственно на веб-сервере. Не забудьте указать URL-адрес в элементе **SourceLocation** файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="965b0-p104">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="965b0-132">В Excel, Word или PowerPoint откройте на ленте вкладку **Вставка** и выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="965b0-132">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="965b0-133">Нажмите **ОБЩАЯ ПАПКА** в верхней части диалогового окна **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="965b0-133">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="965b0-134">Выберите имя надстройки и нажмите кнопку **ОК**, чтобы вставить надстройку.</span><span class="sxs-lookup"><span data-stu-id="965b0-134">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="965b0-135">См. также</span><span class="sxs-lookup"><span data-stu-id="965b0-135">See also</span></span>

- [<span data-ttu-id="965b0-136">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="965b0-136">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="965b0-137">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="965b0-137">Publish your Office Add-in</span></span>](../publish/publish.md)
    
