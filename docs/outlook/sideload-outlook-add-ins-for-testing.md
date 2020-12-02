---
title: Загрузка неопубликованных надстроек Outlook для тестирования
description: Используйте загрузку неопубликованных надстроек, чтобы установить надстройку Outlook для тестирования, не размещая ее в каталоге надстроек.
ms.date: 12/01/2020
localization_priority: Normal
ms.openlocfilehash: dea2125ccd64eba2e3f1695c8ca1111a710321a4
ms.sourcegitcommit: c2fd7f982f3da748ef6be5c3a7434d859f8b46b9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/02/2020
ms.locfileid: "49530929"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="37fbf-103">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="37fbf-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="37fbf-104">Вы можете использовать загрузку неопубликованных надстроек, чтобы установить надстройку Outlook для тестирования, не размещая ее в каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="37fbf-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a><span data-ttu-id="37fbf-105">Загрузка неопубликованной надстройки в Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="37fbf-105">Sideload an add-in in Outlook on the web</span></span>

<span data-ttu-id="37fbf-106">Процесс загрузки неопубликованной надстройки в Outlook в Интернете зависит от того, используется ли новая или классическая версия.</span><span class="sxs-lookup"><span data-stu-id="37fbf-106">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="37fbf-107">Если ваша панель инструментов почтового ящика выглядит так, как показано на изображении ниже, см. статью [Загрузка неопубликованных надстроек в новой веб-версии Outlook](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="37fbf-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![снимок части экрана с изображением веб-панели инструментов новой веб-версии Outlook](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="37fbf-109">Если ваша панель инструментов почтового ящика выглядит так, как показано на изображении ниже, см. статью [Загрузка неопубликованных надстроек в классической веб-версии Outlook](#sideload-an-add-in-in-classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="37fbf-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![снимок части экрана с изображением веб-панели инструментов классической веб-версии Outlook](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="37fbf-111">Если ваша организация добавили свой логотип на панель инструментов почтового ящика, вы можете увидеть изображение, которое будет немного отличаться от показанных ранее изображений.</span><span class="sxs-lookup"><span data-stu-id="37fbf-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="37fbf-112">Загрузка неопубликованной надстройки в новой веб-версии Outlook</span><span class="sxs-lookup"><span data-stu-id="37fbf-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="37fbf-113">Перейдите к [Outlook в Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="37fbf-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="37fbf-114">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="37fbf-114">In Outlook on the web, create a new message.</span></span>

1. <span data-ttu-id="37fbf-115">Выберите **...** в нижней части нового сообщения, а затем выберите **Получить надстройки** в появившемся меню.</span><span class="sxs-lookup"><span data-stu-id="37fbf-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Окно создания сообщений в новой веб-версии Outlook с выделенной опцией "Получить надстройки"](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="37fbf-117">В диалоговом окне **Надстройки для Outlook** выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Диалоговое окно "Надстройки для Outlook" в новой веб-версии Outlook с выбранной опцией "Мои надстройки "](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="37fbf-119">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="37fbf-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="37fbf-120">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: управление надстройками с указанием параметра "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="37fbf-p102">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="37fbf-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="37fbf-124">Загрузка неопубликованной надстройки в классической веб-версии Outlook</span><span class="sxs-lookup"><span data-stu-id="37fbf-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="37fbf-125">Перейдите к [Outlook в Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="37fbf-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="37fbf-126">Нажмите значок шестеренки в верхнем правом углу панели инструментов и выберите пункт **Управление надстройками**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Снимок экрана: веб-версия Outlook с параметром "Управление надстройками"](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="37fbf-128">На странице **Управление надстройками** выберите **Надстройки** > **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Диалоговое окно магазина веб-версии Outlook с открытым разделом "Мои надстройки"](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="37fbf-130">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="37fbf-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="37fbf-131">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: управление надстройками с указанием параметра "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="37fbf-p104">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="37fbf-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="37fbf-135">Загрузка неопубликованной надстройки в классической версии Outlook</span><span class="sxs-lookup"><span data-stu-id="37fbf-135">Sideload an add-in in Outlook on the desktop</span></span>

### <a name="outlook-2016-or-later"></a><span data-ttu-id="37fbf-136">Outlook 2016 или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="37fbf-136">Outlook 2016 or later</span></span>

1. <span data-ttu-id="37fbf-137">Откройте Outlook 2016 или более поздней версии в Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="37fbf-137">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="37fbf-138">Нажмите кнопку **Получить надстройки** на ленте.</span><span class="sxs-lookup"><span data-stu-id="37fbf-138">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Лента Outlook 2016, указывающая на кнопку получения надстроек](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="37fbf-140">Если вы не видите кнопку **получить** надстройки в вашей версии Outlook, выберите:</span><span class="sxs-lookup"><span data-stu-id="37fbf-140">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="37fbf-141">Кнопка " **сохранить** " на ленте (если она доступна).</span><span class="sxs-lookup"><span data-stu-id="37fbf-141">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="37fbf-142">OR</span><span class="sxs-lookup"><span data-stu-id="37fbf-142">OR</span></span>
    >
    > - <span data-ttu-id="37fbf-143">Меню **файл** , а затем нажмите кнопку **Управление надстройками** на вкладке **сведения** , чтобы открыть диалоговое окно **надстройки** в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="37fbf-143">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="37fbf-144">Более подробную информацию о веб-интерфейсе можно узнать в предыдущем разделе [Загрузка неопубликованных надстройки в Outlook в Интернете](#sideload-an-add-in-in-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="37fbf-144">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web).</span></span>

1. <span data-ttu-id="37fbf-145">Если в верхней части диалогового окна есть вкладки, убедитесь, что выбрана вкладка **надстройки** .</span><span class="sxs-lookup"><span data-stu-id="37fbf-145">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="37fbf-146">Выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-146">Choose **My add-ins**.</span></span>

    ![Диалоговое окно магазина Outlook 2016 с открытым разделом "Мои надстройки"](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="37fbf-148">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="37fbf-148">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="37fbf-149">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="37fbf-149">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: магазин с параметром "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="37fbf-p107">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="37fbf-p107">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-2013"></a><span data-ttu-id="37fbf-153">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="37fbf-153">Outlook 2013</span></span>

1. <span data-ttu-id="37fbf-154">Откройте Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="37fbf-154">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="37fbf-155">Выберите меню **файл** , а затем нажмите кнопку **Управление надстройками** на вкладке **сведения** . Outlook откроет веб-версию в браузере.</span><span class="sxs-lookup"><span data-stu-id="37fbf-155">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="37fbf-156">Выполните действия, описанные в разделе [Загрузка неопубликованных надстройка в Outlook в Интернете,](#sideload-an-add-in-in-outlook-on-the-web) в соответствии с вашей версией Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="37fbf-156">Follow the steps in the [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="37fbf-157">Удаление надстройки неопубликованные</span><span class="sxs-lookup"><span data-stu-id="37fbf-157">Remove a sideloaded add-in</span></span>

<span data-ttu-id="37fbf-158">Чтобы удалить надстройку неопубликованные из Outlook, выполните действия, описанные ранее в этой статье, чтобы найти надстройку в разделе **Настраиваемые** надстройки диалогового окна со списком установленных надстроек. Нажмите кнопку с многоточием ( `...` ) для надстройки, а затем нажмите кнопку **Удалить** , чтобы удалить эту надстройку.</span><span class="sxs-lookup"><span data-stu-id="37fbf-158">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>