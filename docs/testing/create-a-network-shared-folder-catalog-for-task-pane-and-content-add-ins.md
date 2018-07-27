---
title: Загрузка неопубликованных надстроек Office для тестирования
description: ''
ms.date: 01/25/2018
ms.openlocfilehash: e5769ef40868ec996194725d98913e61b76279bc
ms.sourcegitcommit: 9e0952b3df852bd2896e9f4a6f59f5b89fc1ae24
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/27/2018
ms.locfileid: "21270295"
---
# <a name="sideload-office-add-ins-for-testing"></a><span data-ttu-id="17bcd-102">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="17bcd-102">Sideload Office Add-ins for testing</span></span>

<span data-ttu-id="17bcd-103">Вы можете установить надстройку Office для тестирования в клиенте Office, запущенном в Windows одним из следующих способов:</span><span class="sxs-lookup"><span data-stu-id="17bcd-103">You can install an Office Add-in for testing in an Office client running on Windows by one of the following methods:</span></span>

- <span data-ttu-id="17bcd-104">Использование каталога общих папок для публикации манифеста в общем файловом ресурсе (инструкции ниже)</span><span class="sxs-lookup"><span data-stu-id="17bcd-104">Using a shared folder catalog to publish the manifest to a network file share (instructions below)</span></span>
- [<span data-ttu-id="17bcd-105">Запуск команды"**npm run sideload**" из корня папки проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="17bcd-105">Running the "**npm run sideload**" command from the root of the add-in project folder.</span></span>](sideload-office-addin-using-sideload-command.md) 
>[!NOTE]
><span data-ttu-id="17bcd-106">Метод "npm run sideload" работает только для надстроек Excel, Word и PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="17bcd-106">The "npm run sideload" method only works for Excel, Word, and PowerPoint add-ins).</span></span>

<span data-ttu-id="17bcd-107">Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей:</span><span class="sxs-lookup"><span data-stu-id="17bcd-107">If you're not testing a Word, Excel, or PowerPoint add-in on Windows, see one of the following topics to sideload your add-in:</span></span>

- [<span data-ttu-id="17bcd-108">Загрузка неопубликованных надстроек Office в Office Online для тестирования</span><span class="sxs-lookup"><span data-stu-id="17bcd-108">Sideload Office Add-ins in Office Online for testing</span></span>](sideload-office-add-ins-for-testing.md)
- [<span data-ttu-id="17bcd-109">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="17bcd-109">Sideload Office Add-ins on iPad and Mac for testing</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)

<span data-ttu-id="17bcd-110">В следующем видео вы узнаете ознакомитесь с процессом загрузки вашей надстройки на рабочий стол Office или Office Online с помощью каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="17bcd-110">The following video walks you through the process of sideloading your add-in on Office desktop or Office Online.</span></span>  


> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]


## <a name="share-a-folder"></a><span data-ttu-id="17bcd-111">Общий доступ к папке</span><span class="sxs-lookup"><span data-stu-id="17bcd-111">Share a folder</span></span>

1. <span data-ttu-id="17bcd-112">На том компьютере с Windows, где должна размещаться надстройка, перейдите к родительской папке или диску с папкой, которую требуется использовать в качестве каталога общих папок.</span><span class="sxs-lookup"><span data-stu-id="17bcd-112">On the Windows computer where you want to host your add-in, go to the parent folder, or drive letter, of the folder you want to use as your shared folder catalog.</span></span>

2. <span data-ttu-id="17bcd-113">Откройте контекстное меню папки (щелкните ее правой кнопкой мыши) и выберите пункт **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-113">Open the context menu for the folder (right-click) and choose **Properties**.</span></span>

3. <span data-ttu-id="17bcd-114">Откройте вкладку **Доступ**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-114">Open the **Sharing** tab.</span></span>

4. <span data-ttu-id="17bcd-p101">На странице **Выбор людей** добавьте себя и других пользователей, которым требуется предоставить доступ к надстройке. Если все эти пользователи являются участниками группы безопасности, вы можете добавить группу. Вам потребуются разрешения на **чтение и запись** папки.</span><span class="sxs-lookup"><span data-stu-id="17bcd-p101">On the **Choose people ...** page, add yourself and and anyone else with whom you want to share your add-in. If they are all members of a security group, you can add the group. You will need at least **Read/Write** permission to the folder.</span></span> 

5. <span data-ttu-id="17bcd-118">Нажмите кнопки **Общий доступ** > **Готово** > **Закрыть**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-118">Choose **Share** > **Done** > **Close**.</span></span>


## <a name="specify-the-shared-folder-as-a-trusted-catalog"></a><span data-ttu-id="17bcd-119">Укажите общую папку в качестве каталога общих папок</span><span class="sxs-lookup"><span data-stu-id="17bcd-119">Specify the shared folder as a trusted catalog</span></span>
      
1. <span data-ttu-id="17bcd-120">Откройте новый документ в Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="17bcd-120">Open a new document in Excel, Word, or PowerPoint.</span></span>
    
2. <span data-ttu-id="17bcd-121">Перейдите на вкладку **Файл**, а затем выберите **Параметры**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-121">Choose the **File** tab, and then choose **Options**.</span></span>
    
3. <span data-ttu-id="17bcd-122">Выберите **Центр управления безопасностью**, а затем нажмите кнопку **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-122">Choose **Trust Center**, and then choose the  **Trust Center Settings** button.</span></span>
    
4. <span data-ttu-id="17bcd-123">Выберите пункт **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-123">Choose  **Trusted Add-in Catalogs**.</span></span>
    
5. <span data-ttu-id="17bcd-124">В поле **URL-адрес каталога** введите полный сетевой путь к каталогу общих папок и нажмите **Добавить каталог**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-124">In the  **Catalog Url** box, enter the full network path to the shared folder catalog, and then choose **Add Catalog**.</span></span>
    
6. <span data-ttu-id="17bcd-125">Установите флажок **Показать в меню** и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-125">Select the **Show in Menu** check box, and then choose **OK**.</span></span>

7. <span data-ttu-id="17bcd-126">Закройте приложение Office, чтобы изменения вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="17bcd-126">Close the Office application so your changes will take effect.</span></span>
    

## <a name="sideload-your-add-in"></a><span data-ttu-id="17bcd-127">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="17bcd-127">Sideload your add-in</span></span>

1. <span data-ttu-id="17bcd-p102">Файл манифеста тестируемой надстройки необходимо поместить в каталог общих папок. Обратите внимание, что вы развертываете веб-приложение непосредственно на веб-сервере. Не забудьте указать URL-адрес в элементе **SourceLocation** файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="17bcd-p102">Put the manifest file of any add-in that you are testing in the shared folder catalog. Note that you deploy the web application itself to a web server. Be sure to specify the URL in the **SourceLocation** element of the manifest file.</span></span>

    > [!IMPORTANT]
    > [!include[HTTPS guidance](../includes/https-guidance.md)]

2. <span data-ttu-id="17bcd-131">В Excel, Word или PowerPoint откройте на ленте вкладку **Вставка** и выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-131">In Excel, Word, or PowerPoint, select **My Add-ins** on the **Insert** tab of the ribbon.</span></span>

3. <span data-ttu-id="17bcd-132">Нажмите **ОБЩАЯ ПАПКА** в верхней части диалогового окна **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="17bcd-132">Choose **SHARED FOLDER** at the top of the **Office Add-ins** dialog box.</span></span>

4. <span data-ttu-id="17bcd-133">Выберите имя надстройки и нажмите кнопку **ОК**, чтобы вставить надстройку.</span><span class="sxs-lookup"><span data-stu-id="17bcd-133">Select the name of the add-in and choose **OK** to insert the add-in.</span></span>


## <a name="see-also"></a><span data-ttu-id="17bcd-134">См. также</span><span class="sxs-lookup"><span data-stu-id="17bcd-134">See also</span></span>

- [<span data-ttu-id="17bcd-135">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="17bcd-135">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="17bcd-136">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="17bcd-136">Publish your Office Add-in</span></span>](../publish/publish.md)
    
