---
title: Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования
description: Проверьте Office надстройку на iPad Mac с помощью боковой загрузки.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: b3d7d7fa3ee69e849c112c888b66fa9deed23d88
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076205"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="e75c8-103">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="e75c8-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="e75c8-p101">Чтобы проверить работу надстройки в Office для iOS, вы можете загрузить манифест неопубликованной надстройки на iPad с помощью iTunes или непосредственно в Office для Mac. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать.</span><span class="sxs-lookup"><span data-stu-id="e75c8-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="e75c8-106">Предварительные требования (Office для iOS)</span><span class="sxs-lookup"><span data-stu-id="e75c8-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="e75c8-107">Компьютер Windows или Mac, на котором установлено приложение [iTunes](https://www.apple.com/itunes/download/).</span><span class="sxs-lookup"><span data-stu-id="e75c8-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
  > [!IMPORTANT]
  > <span data-ttu-id="e75c8-108">Если вы используете macOS Catalina, [iTunes](https://support.apple.com/HT210200) больше не доступен, поэтому следует следовать инструкциям в разделе Sideload надстройки на Excel или Word на iPad с помощью [macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) позже в этой статье.</span><span class="sxs-lookup"><span data-stu-id="e75c8-108">If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/HT210200) so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.</span></span>

- <span data-ttu-id="e75c8-109">Установлен iPad iOS 8.2 или более [](https://apps.apple.com/app/microsoft-excel/id586683407) поздней Excel [или Word,](https://apps.apple.com/app/microsoft-word/id586447913) а также синхронизированный кабель.</span><span class="sxs-lookup"><span data-stu-id="e75c8-109">An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.</span></span>

- <span data-ttu-id="e75c8-110">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="e75c8-110">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="e75c8-111">Предварительные требования (Office для Mac)</span><span class="sxs-lookup"><span data-stu-id="e75c8-111">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="e75c8-112">Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).</span><span class="sxs-lookup"><span data-stu-id="e75c8-112">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="e75c8-113">Word для Mac версии 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="e75c8-113">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="e75c8-114">Excel для Mac версии 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="e75c8-114">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="e75c8-115">PowerPoint для Mac версии 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="e75c8-115">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="e75c8-116">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="e75c8-116">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a><span data-ttu-id="e75c8-117">Sideload an add-in on Excel word on iPad using iTunes</span><span class="sxs-lookup"><span data-stu-id="e75c8-117">Sideload an add-in on Excel or Word on iPad using iTunes</span></span>

1. <span data-ttu-id="e75c8-118">Подключите iPad к компьютеру с помощью кабеля для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="e75c8-118">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="e75c8-119">Если вы впервые подключите iPad к компьютеру, вам будет предложено использовать **Trust This Computer?**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-119">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="e75c8-120">Выберите **Доверять**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-120">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="e75c8-121">В iTunes под строкой меню выберите значок **iPad**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-121">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="e75c8-122">В левой части iTunes в разделе **Параметры** выберите **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-122">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="e75c8-123">В правой части iTunes прокрутите окно вниз до раздела **Общий доступ к файлам**, а затем в столбце **Надстройки** выберите **Excel** или **Word**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-123">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="e75c8-124">В нижней части **столбца Excel** **или Word Documents** выберите Добавить **файл,** а затем выберите файл манифеста .xml надстройки, необходимой для загрузки.</span><span class="sxs-lookup"><span data-stu-id="e75c8-124">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="e75c8-125">Откройте приложение Excel или Word на iPad.</span><span class="sxs-lookup"><span data-stu-id="e75c8-125">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="e75c8-126">Если приложение Excel Word уже запущено, выберите кнопку **Главная,** а затем закрой и перезапустите приложение.</span><span class="sxs-lookup"><span data-stu-id="e75c8-126">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="e75c8-127">Откройте документ.</span><span class="sxs-lookup"><span data-stu-id="e75c8-127">Open a document.</span></span>

8. <span data-ttu-id="e75c8-128">Выберите **надстройки** на вкладке **Вставка.** (На вкладке **Вставить** может потребоваться прокрутка по горизонтали, пока не увидите кнопку **Надстройки.)** Ваша надстройка с боковой загрузкой доступна для вставки под заголовком **Developer** в пользовательском интерфейсе **надстройки.**</span><span class="sxs-lookup"><span data-stu-id="e75c8-128">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Вставьте надстройки в Excel приложении.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a><span data-ttu-id="e75c8-130">Sideload an add-in on Excel word on iPad using macOS Catalina</span><span class="sxs-lookup"><span data-stu-id="e75c8-130">Sideload an add-in on Excel or Word on iPad using macOS Catalina</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e75c8-131">С введением macOS Catalina Apple прекратила [iTunes](https://support.apple.com/HT210200) на Mac и интегрированные функции, необходимые для загрузки приложений в **Finder**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-131">With the introduction of macOS Catalina, [Apple discontinued iTunes on Mac](https://support.apple.com/HT210200) and integrated functionality required to sideload apps into **Finder**.</span></span>

1. <span data-ttu-id="e75c8-132">Подключите iPad к компьютеру с помощью кабеля для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="e75c8-132">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="e75c8-133">Если вы впервые подключите iPad к компьютеру, вам будет предложено использовать **Trust This Computer?**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-133">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="e75c8-134">Выберите **Доверять**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-134">Choose **Trust** to continue.</span></span> <span data-ttu-id="e75c8-135">Вы также можете быть заданы вопросы, если это новый iPad или если вы восстанавливаете один.</span><span class="sxs-lookup"><span data-stu-id="e75c8-135">You may also be asked if this is a new iPad or if you're restoring one.</span></span>

2. <span data-ttu-id="e75c8-136">В Finder в **статье Locations** выберите значок **iPad** ниже панели меню.</span><span class="sxs-lookup"><span data-stu-id="e75c8-136">In Finder, under **Locations**, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="e75c8-137">В верхней части окна Finder нажмите кнопку **Файлы,** а затем **найдите** Excel **или Word**.</span><span class="sxs-lookup"><span data-stu-id="e75c8-137">On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.</span></span>

4. <span data-ttu-id="e75c8-138">Из другого окна Finder перетащите и manifest.xml файл надстройки, который необходимо загрузить в файл **Excel** **Word** в первом окне Finder.</span><span class="sxs-lookup"><span data-stu-id="e75c8-138">From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.</span></span>

5. <span data-ttu-id="e75c8-139">Откройте приложение Excel или Word на iPad.</span><span class="sxs-lookup"><span data-stu-id="e75c8-139">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="e75c8-140">Если приложение Excel Word уже запущено, выберите кнопку **Главная,** а затем закрой и перезапустите приложение.</span><span class="sxs-lookup"><span data-stu-id="e75c8-140">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

6. <span data-ttu-id="e75c8-141">Откройте документ.</span><span class="sxs-lookup"><span data-stu-id="e75c8-141">Open a document.</span></span>

7. <span data-ttu-id="e75c8-142">Выберите **надстройки** на вкладке **Вставка.** (На вкладке **Вставить** может потребоваться прокрутка по горизонтали, пока не увидите кнопку **Надстройки.)** Ваша надстройка с боковой загрузкой доступна для вставки под заголовком **Developer** в пользовательском интерфейсе **надстройки.**</span><span class="sxs-lookup"><span data-stu-id="e75c8-142">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Вставьте надстройки в Excel приложении.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="e75c8-144">Загрузка неопубликованной надстройки в Office для Mac</span><span class="sxs-lookup"><span data-stu-id="e75c8-144">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="e75c8-145">Сведения о загрузке неопубликованной надстройки Outlook для Mac см. в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="e75c8-145">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md).</span></span>

1. <span data-ttu-id="e75c8-146">Откройте **терминал** и перейдите в одну из следующих папок, где вы сохраните файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="e75c8-146">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="e75c8-147">Если папки `wef` нет на компьютере, создайте ее.</span><span class="sxs-lookup"><span data-stu-id="e75c8-147">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="e75c8-148">Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="e75c8-148">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>
    - <span data-ttu-id="e75c8-149">Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="e75c8-149">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="e75c8-150">Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="e75c8-150">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="e75c8-151">Откройте папку в **Finder** с помощью команды `open .` (включая период или точку).</span><span class="sxs-lookup"><span data-stu-id="e75c8-151">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="e75c8-152">Скопируйте файл манифеста надстройки в эту папку.</span><span class="sxs-lookup"><span data-stu-id="e75c8-152">Copy your add-in's manifest file to this folder.</span></span>

    ![Папка Wef в Office на Mac.](../images/all-my-files.png)

3. <span data-ttu-id="e75c8-p108">Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.</span><span class="sxs-lookup"><span data-stu-id="e75c8-p108">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="e75c8-156">В Word выберите **вставьте** надстройки Мои надстройки (выпадаемое меню), а затем  >    >   выберите надстройки.</span><span class="sxs-lookup"><span data-stu-id="e75c8-156">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Мои надстройки в Office на Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="e75c8-p109">Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню.</span><span class="sxs-lookup"><span data-stu-id="e75c8-p109">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="e75c8-161">Проверьте, отображается ли ваша надстройка в Word.</span><span class="sxs-lookup"><span data-stu-id="e75c8-161">Verify that your add-in is displayed in Word.</span></span>

    ![Office Надстройка, отображаемая в Office mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="e75c8-163">Удаление боковой надстройки</span><span class="sxs-lookup"><span data-stu-id="e75c8-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="e75c8-164">Вы можете удалить ранее загруженную надстройку, очищая кэш Office на компьютере.</span><span class="sxs-lookup"><span data-stu-id="e75c8-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="e75c8-165">Сведения о том, как очистить кэш для каждой платформы и приложения, можно найти в статье [Clear the Office кэш.](clear-cache.md)</span><span class="sxs-lookup"><span data-stu-id="e75c8-165">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="e75c8-166">См. также</span><span class="sxs-lookup"><span data-stu-id="e75c8-166">See also</span></span>

- [<span data-ttu-id="e75c8-167">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="e75c8-167">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
