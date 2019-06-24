---
title: Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: b97da347064fdceda847403785e0bf792846dfaa
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126969"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="58606-102">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="58606-102">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="58606-p101">Чтобы проверить работу надстройки в Office для iOS, вы можете загрузить манифест неопубликованной надстройки на iPad с помощью iTunes или непосредственно в Office для Mac. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать.</span><span class="sxs-lookup"><span data-stu-id="58606-p101">To see how your add-in will run in Office for iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office for Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span> 

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="58606-105">Предварительные требования (Office для iOS)</span><span class="sxs-lookup"><span data-stu-id="58606-105">Prerequisites for Office for iOS</span></span>

- <span data-ttu-id="58606-106">Компьютер Windows или Mac, на котором установлено приложение [iTunes](https://www.apple.com/itunes/download/).</span><span class="sxs-lookup"><span data-stu-id="58606-106">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
    
- <span data-ttu-id="58606-107">iPad под управлением iOS 8.2 или более поздней версии, на котором установлено приложение [Excel на iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) и к которому подключен кабель для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="58606-107">An iPad running iOS 8.2 or later with [Excel for iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>
    
- <span data-ttu-id="58606-108">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="58606-108">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="58606-109">Предварительные требования (Office для Mac)</span><span class="sxs-lookup"><span data-stu-id="58606-109">Prerequisites for Office for Mac</span></span>

- <span data-ttu-id="58606-110">Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).</span><span class="sxs-lookup"><span data-stu-id="58606-110">A Mac running OS X v10.10 "Yosemite" or later with [Office for Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>
    
- <span data-ttu-id="58606-111">Word для Mac версии 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="58606-111">Word for Mac version 15.18 (160109)</span></span>
   
- <span data-ttu-id="58606-112">Excel для Mac версии 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="58606-112">Excel for Mac version 15.19 (160206)</span></span>

- <span data-ttu-id="58606-113">PowerPoint для Mac версии 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="58606-113">PowerPoint for Mac version 15.24 (160614)</span></span>
    
- <span data-ttu-id="58606-114">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="58606-114">The manifest .xml file for the add-in you want to test.</span></span>
    

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="58606-115">Загрузка неопубликованной надстройки в Excel или Word на iPad</span><span class="sxs-lookup"><span data-stu-id="58606-115">Sideload an add-in on Excel or Word for iPad</span></span>

1. <span data-ttu-id="58606-p102">Подключите iPad к компьютеру с помощью кабеля для синхронизации. Если вы подключаете iPad к компьютеру в первый раз, появится запрос **Доверять этому компьютеру?**. Выберите **Доверять**.</span><span class="sxs-lookup"><span data-stu-id="58606-p102">Use a sync cable to connect your iPad to your computer. If you're connecting the iPad to your computer for the first time, you'll be prompted with  **Trust This Computer?**. Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="58606-119">В iTunes под строкой меню выберите значок **iPad**.</span><span class="sxs-lookup"><span data-stu-id="58606-119">In iTunes, choose the  **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="58606-120">В левой части iTunes в разделе  **Параметры** выберите **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="58606-120">Under  **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="58606-121">В правой части iTunes прокрутите окно вниз до раздела  **Общий доступ к файлам**, а затем в столбце  **Надстройки** выберите **Excel** или **Word**.</span><span class="sxs-lookup"><span data-stu-id="58606-121">On the right side of iTunes, scroll down to  **File Sharing**, and then choose  **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="58606-122">В нижней части столбца  **Документы Excel** или **Документы Word** выберите элемент **Добавить файл**, а затем выберите XML-файл манифеста для надстройки, которую необходимо загрузить.</span><span class="sxs-lookup"><span data-stu-id="58606-122">At the bottom of the  **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span> 
    
6. <span data-ttu-id="58606-p103">Откройте приложение Excel или Word на iPad. Если приложение Excel или Word уже запущено, нажмите кнопку **Главная**, а затем закройте и перезапустите его.</span><span class="sxs-lookup"><span data-stu-id="58606-p103">Open the Excel or Word app on your iPad. If the Excel or Word app is already running, choose the  **Home** button, and then close and restart the app.</span></span>
    
7. <span data-ttu-id="58606-125">Откройте документ.</span><span class="sxs-lookup"><span data-stu-id="58606-125">Open a document.</span></span>
    
8. <span data-ttu-id="58606-126">На вкладке  **Вставка** выберите **Надстройки**. Загруженную надстройка можно добавить в разделе  **Разработчик** в пользовательском интерфейсе **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="58606-126">Choose  **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>
    
    ![Вставка надстроек в приложение Excel](../images/excel-insert-add-in.png)


## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="58606-128">Загрузка неопубликованной надстройки в Office для Mac</span><span class="sxs-lookup"><span data-stu-id="58606-128">Sideload an add-in on Office for Mac</span></span>

> [!NOTE]
> <span data-ttu-id="58606-129">Сведения о загрузке неопубликованной надстройки Outlook для Mac см. в статье [Загрузка неопубликованных надстроек Outlook для тестирования](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="58606-129">To sideload an Outlook add-in, see [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span>

1. <span data-ttu-id="58606-p104">Откройте **Терминал** и перейдите в одну из указанных ниже папок, чтобы сохранить в нее файл манифеста надстройки. Если папки `wef` нет на компьютере, создайте ее.</span><span class="sxs-lookup"><span data-stu-id="58606-p104">Open  **Terminal** and go to one of the following folders where you'll save your add-in's manifest file. If the `wef` folder doesn't exist on your computer, create it.</span></span>
    
    - <span data-ttu-id="58606-132">Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="58606-132">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`</span></span>    
    - <span data-ttu-id="58606-133">Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="58606-133">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`</span></span>
    - <span data-ttu-id="58606-134">Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span><span class="sxs-lookup"><span data-stu-id="58606-134">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`</span></span>
    
2. <span data-ttu-id="58606-p105">Откройте папку в **Finder** с помощью команды `open .` (включая точку). Скопируйте файл манифеста надстройки в эту папку.</span><span class="sxs-lookup"><span data-stu-id="58606-p105">Open the folder in  **Finder** using the command `open .` (including the period or dot). Copy your add-in's manifest file to this folder.</span></span>
    
    ![Папка Wef в Office для Mac](../images/all-my-files.png)

3. <span data-ttu-id="58606-p106">Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.</span><span class="sxs-lookup"><span data-stu-id="58606-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>
    
4. <span data-ttu-id="58606-140">В Word выберите элементы **Вставка**  >  **Надстройки**  >  **Мои надстройки**, а затем выберите свою надстройку.</span><span class="sxs-lookup"><span data-stu-id="58606-140">In Word, choose  **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>
    
    ![Мои надстройки в Office для Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="58606-p107">Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню.</span><span class="sxs-lookup"><span data-stu-id="58606-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span> 
    
5. <span data-ttu-id="58606-145">Проверьте, отображается ли ваша надстройка в Word.</span><span class="sxs-lookup"><span data-stu-id="58606-145">Verify that your add-in is displayed in Word.</span></span>
    
    ![Надстройка в Office для Mac](../images/lorem-ipsum-wikipedia.png)
    
### <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="58606-147">Очистка кэша приложения Office на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="58606-147">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="58606-148">Для повышения производительности надстройки часто кэшируются в Office для Mac.</span><span class="sxs-lookup"><span data-stu-id="58606-148">Add-ins are cached often in Office for Mac, for performance reasons.</span></span> <span data-ttu-id="58606-149">Как правило, для очистки кэша необходимо перезагрузить надстройку.</span><span class="sxs-lookup"><span data-stu-id="58606-149">Normally, the cache is cleared by reloading the add-in.</span></span> <span data-ttu-id="58606-150">Если в одном документе несколько надстроек, автоматическая очистка кэша может не сработать при перезагрузке.</span><span class="sxs-lookup"><span data-stu-id="58606-150">If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="58606-151">На компьютере Mac можно очистить кэш вручную, удалив все содержимое папки `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="58606-151">On a Mac, you can clear the cache manually by deleting everything in the `~/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

## <a name="see-also"></a><span data-ttu-id="58606-152">См. также</span><span class="sxs-lookup"><span data-stu-id="58606-152">See also</span></span>

- [<span data-ttu-id="58606-153">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="58606-153">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
