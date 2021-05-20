---
title: Загрузка неопубликованных надстроек Outlook для тестирования
description: Используйте загрузку неопубликованных надстроек, чтобы установить надстройку Outlook для тестирования, не размещая ее в каталоге надстроек.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 9d0fb246f6522c745658a09fce6934ee44d5079a
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555194"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="a622b-103">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="a622b-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="a622b-104">Вы можете использовать загрузку неопубликованных надстроек, чтобы установить надстройку Outlook для тестирования, не размещая ее в каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="a622b-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="a622b-105">Боковая нагрузка автоматически</span><span class="sxs-lookup"><span data-stu-id="a622b-105">Sideload automatically</span></span>

<span data-ttu-id="a622b-106">Если вы создали Outlook надстройку с [помощью генератора Yeoman для Office дополнительных надстройок,](https://github.com/OfficeDev/generator-office)то загрузку лучше всего делать через командную строку.</span><span class="sxs-lookup"><span data-stu-id="a622b-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="a622b-107">Это позволит воспользоваться нашими инструментами и боковой загрузкой на всех поддерживаемых устройствах в одной команде.</span><span class="sxs-lookup"><span data-stu-id="a622b-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="a622b-108">Используя командную строку, перейдите к корневому каталогу проекта надстройки, генерируемого Yeoman.</span><span class="sxs-lookup"><span data-stu-id="a622b-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="a622b-109">Выполните команду `npm start`.</span><span class="sxs-lookup"><span data-stu-id="a622b-109">Run the command `npm start`.</span></span>

1. <span data-ttu-id="a622b-110">Ваше Outlook надстройка автоматически будет перегружена для Outlook на вашем настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="a622b-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="a622b-111">Вы увидите, что появляется диалог, заявляя, что есть попытка перезагрузить надстройку, перечислив имя и местоположение файла манифеста.</span><span class="sxs-lookup"><span data-stu-id="a622b-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="a622b-112">Выберите **OK**, который будет регистрировать манифест.</span><span class="sxs-lookup"><span data-stu-id="a622b-112">Select **OK**, which will register the manifest.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="a622b-113">Если манифест содержит ошибку или путь к манифесту недействителен, вы получите сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="a622b-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

1. <span data-ttu-id="a622b-114">Если ваш манифест не содержит ошибок и путь действителен, ваше дополнение теперь будет загружено и доступно как на рабочем столе, так и Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a622b-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="a622b-115">Он также будет установлен на всех поддерживаемых устройствах.</span><span class="sxs-lookup"><span data-stu-id="a622b-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="a622b-116">Боковая нагрузка вручную</span><span class="sxs-lookup"><span data-stu-id="a622b-116">Sideload manually</span></span>

<span data-ttu-id="a622b-117">Хотя мы настоятельно рекомендуем автоматически перегрузить по командной строке, как это было в предыдущем разделе, вы также можете вручную перезагрузить Outlook надстройку на основе Outlook клиента.</span><span class="sxs-lookup"><span data-stu-id="a622b-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="a622b-118">Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="a622b-118">Outlook on the web</span></span>

<span data-ttu-id="a622b-119">Процесс боковой загрузки надстройки в интернете Outlook зависит от того, используете ли вы новую или классическую версию.</span><span class="sxs-lookup"><span data-stu-id="a622b-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="a622b-120">Если ваша панель инструментов почтового ящика выглядит так, как показано на изображении ниже, см. статью [Загрузка неопубликованных надстроек в новой веб-версии Outlook](#new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="a622b-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![снимок части экрана с изображением веб-панели инструментов новой веб-версии Outlook](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="a622b-122">Если ваша панель инструментов почтового ящика выглядит так, как показано на изображении ниже, см. статью [Загрузка неопубликованных надстроек в классической веб-версии Outlook](#classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="a622b-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![снимок части экрана с изображением веб-панели инструментов классической веб-версии Outlook](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="a622b-124">Если ваша организация добавили свой логотип на панель инструментов почтового ящика, вы можете увидеть изображение, которое будет немного отличаться от показанных ранее изображений.</span><span class="sxs-lookup"><span data-stu-id="a622b-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="a622b-125">Новые Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="a622b-125">New Outlook on the web</span></span>

1. <span data-ttu-id="a622b-126">Откройте [Outlook в Интернете](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="a622b-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="a622b-127">Создайте новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="a622b-127">Create a new message.</span></span>

1. <span data-ttu-id="a622b-128">Выберите **...** в нижней части нового сообщения, а затем выберите **Получить надстройки** в появившемся меню.</span><span class="sxs-lookup"><span data-stu-id="a622b-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Окно создания сообщений в новой веб-версии Outlook с выделенной опцией "Получить надстройки"](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="a622b-130">В диалоговом окне **Надстройки для Outlook** выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="a622b-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Диалоговое окно "Надстройки для Outlook" в новой веб-версии Outlook с выбранной опцией "Мои надстройки "](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="a622b-132">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="a622b-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="a622b-133">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="a622b-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: управление надстройками с указанием параметра "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="a622b-p106">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="a622b-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="a622b-137">Классический Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="a622b-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="a622b-138">Откройте [Outlook в Интернете](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="a622b-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="a622b-139">Нажмите значок шестеренки в верхнем правом углу панели инструментов и выберите пункт **Управление надстройками**.</span><span class="sxs-lookup"><span data-stu-id="a622b-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Снимок экрана: веб-версия Outlook с параметром "Управление надстройками"](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="a622b-141">На странице **Управление надстройками** выберите **Надстройки** > **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="a622b-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Диалоговое окно магазина веб-версии Outlook с открытым разделом "Мои надстройки"](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="a622b-143">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="a622b-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="a622b-144">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="a622b-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: управление надстройками с указанием параметра "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="a622b-p108">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="a622b-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="a622b-148">Outlook на рабочем столе</span><span class="sxs-lookup"><span data-stu-id="a622b-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="a622b-149">Outlook 2016 или позже</span><span class="sxs-lookup"><span data-stu-id="a622b-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="a622b-150">Откройте Outlook 2016 или позже на Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="a622b-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="a622b-151">Нажмите кнопку **Получить надстройки** на ленте.</span><span class="sxs-lookup"><span data-stu-id="a622b-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016 лента, указывающая на кнопку «Получить надстройки»](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="a622b-153">Если вы не видите **кнопку Get Add-ins** в версии Outlook, выберите:</span><span class="sxs-lookup"><span data-stu-id="a622b-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="a622b-154">**Храните** кнопку на ленте, если она доступна.</span><span class="sxs-lookup"><span data-stu-id="a622b-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="a622b-155">OR</span><span class="sxs-lookup"><span data-stu-id="a622b-155">OR</span></span>
    >
    > - <span data-ttu-id="a622b-156">**Меню** файлов, а затем **выберите кнопку «Управлять надстройки»** **на вкладке Info,** чтобы открыть **диалог add-ins** Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a622b-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="a622b-157">Вы можете увидеть больше о веб-опыт в предыдущем разделе [Sideload надстройки в Outlook в Интернете](#outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="a622b-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="a622b-158">Если в верхней части диалога есть вкладки, убедитесь, **что вкладка «Дополнительные встрой»** выбрана.</span><span class="sxs-lookup"><span data-stu-id="a622b-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="a622b-159">Выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="a622b-159">Choose **My add-ins**.</span></span>

    ![Диалоговое окно магазина Outlook 2016 с открытым разделом "Мои надстройки"](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="a622b-161">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="a622b-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="a622b-162">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="a622b-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: магазин с параметром "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="a622b-p111">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="a622b-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="a622b-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="a622b-166">Outlook 2013</span></span>

1. <span data-ttu-id="a622b-167">Открыт Outlook 2013 года в Windows.</span><span class="sxs-lookup"><span data-stu-id="a622b-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="a622b-168">Выберите **меню** файла, а затем выберите кнопку «Управлять надстройки» на вкладке **Info.** Outlook она откроет **веб-версию** в браузере.</span><span class="sxs-lookup"><span data-stu-id="a622b-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="a622b-169">Следуйте шагам в [Sideload надстройки в разделе Outlook в соответствии](#outlook-on-the-web) с вашей версией Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="a622b-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="a622b-170">Удалите боковое надстройку</span><span class="sxs-lookup"><span data-stu-id="a622b-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="a622b-171">На всех версиях Outlook, ключом к удалению боковой загрузки надстройки **является диалог My Add-ins,** в котором перечислены установленные надстройки. Выберите эллипсис `...` () для дополнения, а затем выберите **Удалить**.</span><span class="sxs-lookup"><span data-stu-id="a622b-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="a622b-172">Чтобы перейти к **диалоговому ящику My Add-ins** для вашего клиента Outlook, используйте последние шаги, перечисленные [для ручной боковой загрузки](#sideload-manually) в предыдущих разделах этой статьи.</span><span class="sxs-lookup"><span data-stu-id="a622b-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="a622b-173">Чтобы удалить боковое дополнение с Outlook, используйте шаги, описанные ранее в этой статье, чтобы найти надстройку **в разделе пользовательских надстройок диалогового** окна, в которую перечислены установленные надстройки. Выберите эллипсис `...` () для дополнения, а затем выберите **Удалить,** чтобы удалить это конкретное дополнение.</span><span class="sxs-lookup"><span data-stu-id="a622b-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="a622b-174">Закройте диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="a622b-174">Close the dialog.</span></span>
