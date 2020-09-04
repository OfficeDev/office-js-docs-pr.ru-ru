---
title: Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования
description: Протестируйте надстройку Office на iPad и Mac, не прополнив загрузку неопубликованных приложений.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7c5e9542c6e6f9abc96defde389b9543421b8529
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/03/2020
ms.locfileid: "47364060"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a><span data-ttu-id="bdb99-103">Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования</span><span class="sxs-lookup"><span data-stu-id="bdb99-103">Sideload Office Add-ins on iPad and Mac for testing</span></span>

<span data-ttu-id="bdb99-p101">Чтобы проверить работу надстройки в Office для iOS, вы можете загрузить манифест неопубликованной надстройки на iPad с помощью iTunes или непосредственно в Office для Mac. Вы не сможете устанавливать точки останова и отлаживать код надстройки во время выполнения, но сможете проверить ее работу и убедиться, что интерфейс отображается правильно и его можно использовать.</span><span class="sxs-lookup"><span data-stu-id="bdb99-p101">To see how your add-in will run in Office on iOS, you can sideload your add-in's manifest onto an iPad using iTunes, or sideload your add-in's manifest directly in Office on Mac. This action won't enable you to set breakpoints and debug your add-in's code while it's running, but you can see how it behaves and verify that the UI is usable and rendering appropriately.</span></span>

## <a name="prerequisites-for-office-on-ios"></a><span data-ttu-id="bdb99-106">Предварительные требования (Office для iOS)</span><span class="sxs-lookup"><span data-stu-id="bdb99-106">Prerequisites for Office on iOS</span></span>

- <span data-ttu-id="bdb99-107">Компьютер Windows или Mac, на котором установлено приложение [iTunes](https://www.apple.com/itunes/download/).</span><span class="sxs-lookup"><span data-stu-id="bdb99-107">A Windows or Mac computer with [iTunes](https://www.apple.com/itunes/download/) installed.</span></span>
  > [!IMPORTANT]
  > <span data-ttu-id="bdb99-108">Если вы используете macOS Каталина, [iTunes больше не будет доступен](https://support.apple.com/HT210200) , поэтому вам следует выполнить инструкции из раздела [Загрузка неопубликованных надстройки в Excel или Word на iPad с помощью macOS Каталина](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="bdb99-108">If you're running macOS Catalina, [iTunes is no longer available](https://support.apple.com/HT210200) so you should follow the instructions in the section [Sideload an add-in on Excel or Word on iPad using macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) later in this article.</span></span>

- <span data-ttu-id="bdb99-109">IPad под управлением iOS 8,2 или более поздней версии с установленными [приложением Excel](https://apps.apple.com/app/microsoft-excel/id586683407) или [Word](https://apps.apple.com/app/microsoft-word/id586447913) и кабелем синхронизации.</span><span class="sxs-lookup"><span data-stu-id="bdb99-109">An iPad running iOS 8.2 or later with [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) or [Word](https://apps.apple.com/app/microsoft-word/id586447913) installed, and a sync cable.</span></span>

- <span data-ttu-id="bdb99-110">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="bdb99-110">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="prerequisites-for-office-on-mac"></a><span data-ttu-id="bdb99-111">Предварительные требования (Office для Mac)</span><span class="sxs-lookup"><span data-stu-id="bdb99-111">Prerequisites for Office on Mac</span></span>

- <span data-ttu-id="bdb99-112">Компьютер Mac под управлением OS X 10.10 Yosemite или более поздней версии с установленным набором [Office для Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac).</span><span class="sxs-lookup"><span data-stu-id="bdb99-112">A Mac running OS X v10.10 "Yosemite" or later with [Office on Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installed.</span></span>

- <span data-ttu-id="bdb99-113">Word для Mac версии 15.18 (160109).</span><span class="sxs-lookup"><span data-stu-id="bdb99-113">Word on Mac version 15.18 (160109).</span></span>

- <span data-ttu-id="bdb99-114">Excel для Mac версии 15.19 (160206).</span><span class="sxs-lookup"><span data-stu-id="bdb99-114">Excel on Mac version 15.19 (160206).</span></span>

- <span data-ttu-id="bdb99-115">PowerPoint для Mac версии 15.24 (160614)</span><span class="sxs-lookup"><span data-stu-id="bdb99-115">PowerPoint on Mac version 15.24 (160614)</span></span>

- <span data-ttu-id="bdb99-116">XML-файл манифеста для надстройки, которую вы хотите протестировать.</span><span class="sxs-lookup"><span data-stu-id="bdb99-116">The manifest .xml file for the add-in you want to test.</span></span>

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a><span data-ttu-id="bdb99-117">Загрузка неопубликованных надстройку в Excel или Word на iPad с помощью iTunes</span><span class="sxs-lookup"><span data-stu-id="bdb99-117">Sideload an add-in on Excel or Word on iPad using iTunes</span></span>

1. <span data-ttu-id="bdb99-118">Подключите iPad к компьютеру с помощью кабеля для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="bdb99-118">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="bdb99-119">Если вы подключаете iPad к компьютеру в первый раз, вам будет предложено **доверять этому компьютеру?**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-119">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="bdb99-120">Выберите **Доверять**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-120">Choose **Trust** to continue.</span></span>

2. <span data-ttu-id="bdb99-121">В iTunes под строкой меню выберите значок **iPad**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-121">In iTunes, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="bdb99-122">В левой части iTunes в разделе **Параметры** выберите **Приложения**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-122">Under **Settings** on the left side of iTunes, choose **Apps**.</span></span>

4. <span data-ttu-id="bdb99-123">В правой части iTunes прокрутите окно вниз до раздела **Общий доступ к файлам**, а затем в столбце **Надстройки** выберите **Excel** или **Word**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-123">On the right side of iTunes, scroll down to **File Sharing**, and then choose **Excel** or **Word** in the **Add-ins** column.</span></span>

5. <span data-ttu-id="bdb99-124">В нижней части столбца документы **Excel** или **Word** нажмите кнопку **Добавить файл**, а затем выберите файл manifest. XML надстройки, которую необходимо Загрузка неопубликованных.</span><span class="sxs-lookup"><span data-stu-id="bdb99-124">At the bottom of the **Excel** or **Word Documents** column, choose **Add File**, and then select the manifest .xml file of the add-in you want to sideload.</span></span>

6. <span data-ttu-id="bdb99-125">Откройте приложение Excel или Word на iPad.</span><span class="sxs-lookup"><span data-stu-id="bdb99-125">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="bdb99-126">Если приложение Excel или Word уже запущено, нажмите кнопку **домой** , а затем закройте и перезапустите приложение.</span><span class="sxs-lookup"><span data-stu-id="bdb99-126">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

7. <span data-ttu-id="bdb99-127">Откройте документ.</span><span class="sxs-lookup"><span data-stu-id="bdb99-127">Open a document.</span></span>

8. <span data-ttu-id="bdb99-128">Выберите **пункт надстройки** на вкладке **Вставка** . (на вкладке **Вставка** может потребоваться прокрутить страницу по горизонтали, пока не появится **кнопка** надстройка.) Надстройка неопубликованные доступна для вставки под заголовком **разработчик** **в пользовательском интерфейсе надстроек.**</span><span class="sxs-lookup"><span data-stu-id="bdb99-128">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Вставка надстроек в приложение Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a><span data-ttu-id="bdb99-130">Загрузка неопубликованных надстройку в Excel или Word на iPad с помощью macOS Каталина</span><span class="sxs-lookup"><span data-stu-id="bdb99-130">Sideload an add-in on Excel or Word on iPad using macOS Catalina</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bdb99-131">С появлением macOS Каталина, [iTunes по протоколу Apple в Mac](https://support.apple.com/HT210200) и интегрированных функциональных возможностей, необходимых для Загрузка неопубликованных приложений в **Finder**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-131">With the introduction of macOS Catalina, [Apple discontinued iTunes on Mac](https://support.apple.com/HT210200) and integrated functionality required to sideload apps into **Finder**.</span></span>

1. <span data-ttu-id="bdb99-132">Подключите iPad к компьютеру с помощью кабеля для синхронизации.</span><span class="sxs-lookup"><span data-stu-id="bdb99-132">Use a sync cable to connect your iPad to your computer.</span></span> <span data-ttu-id="bdb99-133">Если вы подключаете iPad к компьютеру в первый раз, вам будет предложено **доверять этому компьютеру?**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-133">If you're connecting the iPad to your computer for the first time, you'll be prompted with **Trust This Computer?**.</span></span> <span data-ttu-id="bdb99-134">Выберите **Доверять**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-134">Choose **Trust** to continue.</span></span> <span data-ttu-id="bdb99-135">Кроме того, вам может быть предложено сделать это новым iPad или восстановить его.</span><span class="sxs-lookup"><span data-stu-id="bdb99-135">You may also be asked if this is a new iPad or if you're restoring one.</span></span>

2. <span data-ttu-id="bdb99-136">В области Finder в разделе **расположения**выберите значок **iPad** под строкой меню.</span><span class="sxs-lookup"><span data-stu-id="bdb99-136">In Finder, under **Locations**, choose the **iPad** icon below the menu bar.</span></span>

3. <span data-ttu-id="bdb99-137">В верхней части окна Finder выберите **файлы**, а затем найдите **Excel** или **Word**.</span><span class="sxs-lookup"><span data-stu-id="bdb99-137">On the top of the Finder window, click on **Files**, and then locate **Excel** or **Word**.</span></span>

4. <span data-ttu-id="bdb99-138">В другом окне Finder перетащите manifest.xml файл надстройки, который требуется загрузить, в файл **Excel** или **Word** в первом окне Finder.</span><span class="sxs-lookup"><span data-stu-id="bdb99-138">From a different Finder window, drag and drop the manifest.xml file of the add-in you want to side load onto the **Excel** or **Word** file in the first Finder window.</span></span>

5. <span data-ttu-id="bdb99-139">Откройте приложение Excel или Word на iPad.</span><span class="sxs-lookup"><span data-stu-id="bdb99-139">Open the Excel or Word app on your iPad.</span></span> <span data-ttu-id="bdb99-140">Если приложение Excel или Word уже запущено, нажмите кнопку **домой** , а затем закройте и перезапустите приложение.</span><span class="sxs-lookup"><span data-stu-id="bdb99-140">If the Excel or Word app is already running, choose the **Home** button, and then close and restart the app.</span></span>

6. <span data-ttu-id="bdb99-141">Откройте документ.</span><span class="sxs-lookup"><span data-stu-id="bdb99-141">Open a document.</span></span>

7. <span data-ttu-id="bdb99-142">Выберите **пункт надстройки** на вкладке **Вставка** . (на вкладке **Вставка** может потребоваться прокрутить страницу по горизонтали, пока не появится **кнопка** надстройка.) Надстройка неопубликованные доступна для вставки под заголовком **разработчик** **в пользовательском интерфейсе надстроек.**</span><span class="sxs-lookup"><span data-stu-id="bdb99-142">Choose **Add-ins** on the **Insert** tab. (On the **Insert** tab, you may need to scroll horizontally until you see the **Add-ins** button.) Your sideloaded add-in is available to insert under the **Developer** heading in the **Add-ins** UI.</span></span>

    ![Вставка надстроек в приложение Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a><span data-ttu-id="bdb99-144">Загрузка неопубликованной надстройки в Office для Mac</span><span class="sxs-lookup"><span data-stu-id="bdb99-144">Sideload an add-in in Office on Mac</span></span>

> [!NOTE]
> <span data-ttu-id="bdb99-145">Сведения о загрузке неопубликованной надстройки Outlook для Mac см. в статье [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).</span><span class="sxs-lookup"><span data-stu-id="bdb99-145">To sideload an Outlook add-in on Mac, see [Sideload Outlook add-ins for testing](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).</span></span>

1. <span data-ttu-id="bdb99-146">Откройте **терминал** и перейдите в одну из следующих папок, где будет сохранен файл манифеста надстройки.</span><span class="sxs-lookup"><span data-stu-id="bdb99-146">Open **Terminal** and go to one of the following folders where you'll save your add-in's manifest file.</span></span> <span data-ttu-id="bdb99-147">Если папки `wef` нет на компьютере, создайте ее.</span><span class="sxs-lookup"><span data-stu-id="bdb99-147">If the `wef` folder doesn't exist on your computer, create it.</span></span>

    - <span data-ttu-id="bdb99-148">Для Word: `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="bdb99-148">For Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`</span></span>
    - <span data-ttu-id="bdb99-149">Для Excel: `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="bdb99-149">For Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`</span></span>
    - <span data-ttu-id="bdb99-150">Для PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span><span class="sxs-lookup"><span data-stu-id="bdb99-150">For PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`</span></span>

2. <span data-ttu-id="bdb99-151">Откройте папку в **Finder** с помощью команды `open .` (включая точку или точку).</span><span class="sxs-lookup"><span data-stu-id="bdb99-151">Open the folder in **Finder** using the command `open .` (including the period or dot).</span></span> <span data-ttu-id="bdb99-152">Скопируйте файл манифеста надстройки в эту папку.</span><span class="sxs-lookup"><span data-stu-id="bdb99-152">Copy your add-in's manifest file to this folder.</span></span>

    ![Папка Wef в Office для Mac](../images/all-my-files.png)

3. <span data-ttu-id="bdb99-p108">Запустите Word и откройте документ. Если приложение Word уже запущено, перезапустите его.</span><span class="sxs-lookup"><span data-stu-id="bdb99-p108">Open Word, and then open a document. Restart Word if it's already running.</span></span>

4. <span data-ttu-id="bdb99-156">В Word выберите **Вставить**надстройки в  >  **Add-ins**  >  **папку Мои** надстройки (раскрывающееся меню), а затем выберите свою надстройку.</span><span class="sxs-lookup"><span data-stu-id="bdb99-156">In Word, choose **Insert** > **Add-ins** > **My Add-ins** (drop-down menu), and then choose your add-in.</span></span>

    ![Мои надстройки в Office для Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > <span data-ttu-id="bdb99-p109">Неопубликованные надстройки не отображаются в диалоговом окне "Мои надстройки". Они видны только в раскрывающемся меню (небольшая стрелка вниз справа от кнопки "Мои надстройки" на вкладке **Вставка**). Неопубликованные надстройки перечислены под заголовком **Надстройки для разработчиков** в этом меню.</span><span class="sxs-lookup"><span data-stu-id="bdb99-p109">Sideloaded add-ins will not show up in the My Add-ins dialog box. They are only visible within the drop-down menu (small down-arrow to the right of My Add-ins on the **Insert** tab). Sideloaded add-ins are listed under the **Developer Add-ins** heading in this menu.</span></span>

5. <span data-ttu-id="bdb99-161">Проверьте, отображается ли ваша надстройка в Word.</span><span class="sxs-lookup"><span data-stu-id="bdb99-161">Verify that your add-in is displayed in Word.</span></span>

    ![Надстройка в Office для Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="bdb99-163">Удаление надстройки неопубликованные</span><span class="sxs-lookup"><span data-stu-id="bdb99-163">Remove a sideloaded add-in</span></span>

<span data-ttu-id="bdb99-164">Вы можете удалить ранее созданную надстройку неопубликованные, очистив кэш Office на вашем компьютере.</span><span class="sxs-lookup"><span data-stu-id="bdb99-164">You can remove a previously sideloaded add-in by clearing the Office cache on your computer.</span></span> <span data-ttu-id="bdb99-165">Сведения о том, как очистить кэш для каждой платформы и приложений можно найти в статье [Очистка кэша Office](clear-cache.md).</span><span class="sxs-lookup"><span data-stu-id="bdb99-165">Details on how to clear the cache for each platform and application can be found in the article [Clear the Office cache](clear-cache.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="bdb99-166">См. также</span><span class="sxs-lookup"><span data-stu-id="bdb99-166">See also</span></span>

- [<span data-ttu-id="bdb99-167">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="bdb99-167">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)
