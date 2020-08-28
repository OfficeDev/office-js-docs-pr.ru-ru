---
title: Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования
description: Тестирование надстройки Office на iPad и Mac с помощью неопубликованных приложений
ms.date: 02/18/2020
localization_priority: Normal
ms.openlocfilehash: 1a1cb804a72aa182480d06009cf30b41a37276d2
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292204"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="59fda-103">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="59fda-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="59fda-p101">Чтобы проверить работу надстройки в Office для iOS, вы можете загрузить манифест неопубликованной надстройки на iPad с помощью iTunes или непосредственно в Office для Mac. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать.</span><span class="sxs-lookup"><span data-stu-id="59fda-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="59fda-106">Предварительные требования (Office для iOS)</span><span class="sxs-lookup"><span data-stu-id="59fda-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="59fda-107">Компьютер Windows или Mac, на котором установлено приложение [iTunes](https://www.apple.com/itunes/download/).</span><span class="sxs-lookup"><span data-stu-id="59fda-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>

- <span data-ttu-id="59fda-108">iPad под управлением iOS 8.2 или более поздней версии, на котором установлено приложение [Excel на iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) и к которому подключен кабель для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="59fda-108">An iPad running iOS 8.2 or later with [Excel on iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) installed, and a sync cable.</span></span>

- <span data-ttu-id="59fda-109">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="59fda-109">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="59fda-110">Предварительные требования (Office для Mac)</span><span class="sxs-lookup"><span data-stu-id="59fda-110">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="59fda-111">Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).</span><span class="sxs-lookup"><span data-stu-id="59fda-111">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="59fda-112">Word для Mac версии 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="59fda-112">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="59fda-113">Excel для Mac версии 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="59fda-113">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="59fda-114">PowerPoint для Mac версии 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="59fda-114">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="59fda-115">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="59fda-115">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad"></a><span data-ttu-id="59fda-116">Загрузка неопубликованной надстройки в Excel или Word на iPad</span><span class="sxs-lookup"><span data-stu-id="59fda-116">Sideload an add-in on Excel or Word on iPad</span></span>

1. <span data-ttu-id="59fda-117">Подключите iPad к компьютеру с помощью кабеля для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="59fda-117">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="59fda-118">Если вы подключаете iPad к компьютеру в первый раз, вам будет предложено **доверять этому компьютеру?**.</span><span class="sxs-lookup"><span data-stu-id="59fda-118">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="59fda-119">Выберите **Доверять**.</span><span class="sxs-lookup"><span data-stu-id="59fda-119">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="59fda-120">В iTunes под строкой меню выберите значок **iPad**.</span><span class="sxs-lookup"><span data-stu-id="59fda-120">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="59fda-121">В левой части iTunes в разделе **Параметры** выберите **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="59fda-121">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="59fda-122">В правой части iTunes прокрутите окно вниз до раздела **Общий доступ к файлам**, а затем в столбце **Надстройки** выберите **Excel** или **Word**.</span><span class="sxs-lookup"><span data-stu-id="59fda-122">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="59fda-123">В нижней части столбца документы **Excel** или **Word** нажмите кнопку **Добавить файл**, а затем выберите файл manifest. XML надстройки, которую необходимо Загрузка неопубликованных.</span><span class="sxs-lookup"><span data-stu-id="59fda-123">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="59fda-124">Откройте приложение Excel или Word на iPad.</span><span class="sxs-lookup"><span data-stu-id="59fda-124">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="59fda-125">Если приложение Excel или Word уже запущено, нажмите кнопку **домой** , а затем закройте и перезапустите приложение.</span><span class="sxs-lookup"><span data-stu-id="59fda-125">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="59fda-126">Откройте документ.</span><span class="sxs-lookup"><span data-stu-id="59fda-126">Open a document.</span></span>

8. <span data-ttu-id="59fda-127">Выберите **Add-ins** надстройки на вкладке **Вставка** . Надстройка неопубликованные доступна для вставки под заголовком **разработчик** **в пользовательском интерфейсе надстроек.**</span><span class="sxs-lookup"><span data-stu-id="59fda-127">Choose **Add-ins** on the **Insert** tab. Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Вставка надстроек в приложение Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="59fda-129">Загрузка неопубликованной надстройки в Office для Mac</span><span class="sxs-lookup"><span data-stu-id="59fda-129">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="59fda-130">Сведения о загрузке неопубликованной надстройки Outlook для Mac см. в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="59fda-130">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="59fda-131">Откройте **терминал** и перейдите в одну из следующих папок, где будет сохранен файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="59fda-131">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="59fda-132">Если папки `wef` нет на компьютере, создайте ее.</span><span class="sxs-lookup"><span data-stu-id="59fda-132">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="59fda-133">Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="59fda-133">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>    
    - <span data-ttu-id="59fda-134">Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="59fda-134">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="59fda-135">Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="59fda-135">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="59fda-136">Откройте папку в **Finder** с помощью команды `open .` (включая точку или точку).</span><span class="sxs-lookup"><span data-stu-id="59fda-136">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="59fda-137">Скопируйте файл манифеста надстройки в эту папку.</span><span class="sxs-lookup"><span data-stu-id="59fda-137">Copy your add-in's manifest file to this folder.</span></span>

    ![Папка Wef в Office для Mac](../images/all-my-files.png)

3. <span data-ttu-id="59fda-p106">Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.</span><span class="sxs-lookup"><span data-stu-id="59fda-p106">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="59fda-141">В Word выберите **Вставить**надстройки в  >  **Add-ins**  >  **папку Мои** надстройки (раскрывающееся меню), а затем выберите свою надстройку.</span><span class="sxs-lookup"><span data-stu-id="59fda-141">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Мои надстройки в Office для Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="59fda-p107">Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню.</span><span class="sxs-lookup"><span data-stu-id="59fda-p107">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="59fda-146">Проверьте, отображается ли ваша надстройка в Word.</span><span class="sxs-lookup"><span data-stu-id="59fda-146">Verify that your add-in is displayed in Word.</span></span>

    ![Надстройка в Office для Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="59fda-148">Удаление надстройки неопубликованные</span><span class="sxs-lookup"><span data-stu-id="59fda-148">Remove a sideloaded add-in</span></span>

<span data-ttu-id="59fda-149">Вы можете удалить ранее созданную надстройку неопубликованные, очистив кэш Office на вашем компьютере.</span><span class="sxs-lookup"><span data-stu-id="59fda-149">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="59fda-150">Сведения о том, как очистить кэш для каждой платформы и приложений можно найти в статье [Очистка кэша Office](clear-cache.md).</span><span class="sxs-lookup"><span data-stu-id="59fda-150">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="59fda-151">См. также</span><span class="sxs-lookup"><span data-stu-id="59fda-151">See also</span></span>

- [<span data-ttu-id="59fda-152">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="59fda-152">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
