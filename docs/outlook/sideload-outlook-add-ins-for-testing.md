---
title: Загрузка неопубликованных надстроек Outlook для тестирования
description: Используйте загрузку неопубликованных надстроек, чтобы установить надстройку Outlook для тестирования, не размещая ее в каталоге надстроек.
ms.date: 07/09/2020
localization_priority: Normal
ms.openlocfilehash: 9b44b988ddd6552d5f7d14088a0b6f3ae1e410ed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093884"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="2e6cf-103">Загрузка неопубликованных надстроек Outlook для тестирования</span><span class="sxs-lookup"><span data-stu-id="2e6cf-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="2e6cf-104">Вы можете использовать загрузку неопубликованных надстроек, чтобы установить надстройку Outlook для тестирования, не размещая ее в каталоге надстроек.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a><span data-ttu-id="2e6cf-105">Загрузка неопубликованной надстройки в Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="2e6cf-105">Sideload an add-in in Outlook on the web</span></span>

<span data-ttu-id="2e6cf-106">Процесс загрузки неопубликованной надстройки в Outlook в Интернете зависит от того, используется ли новая или классическая версия.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-106">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="2e6cf-107">Если ваша панель инструментов почтового ящика выглядит так, как показано на изображении ниже, см. статью [Загрузка неопубликованных надстроек в новой веб-версии Outlook](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="2e6cf-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![снимок части экрана с изображением веб-панели инструментов новой веб-версии Outlook](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="2e6cf-109">Если ваша панель инструментов почтового ящика выглядит так, как показано на изображении ниже, см. статью [Загрузка неопубликованных надстроек в классической веб-версии Outlook](#sideload-an-add-in-in-classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="2e6cf-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![снимок части экрана с изображением веб-панели инструментов классической веб-версии Outlook](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="2e6cf-111">Если ваша организация добавили свой логотип на панель инструментов почтового ящика, вы можете увидеть изображение, которое будет немного отличаться от показанных ранее изображений.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="2e6cf-112">Загрузка неопубликованной надстройки в новой веб-версии Outlook</span><span class="sxs-lookup"><span data-stu-id="2e6cf-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="2e6cf-113">Перейдите к [Outlook в Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="2e6cf-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="2e6cf-114">Создайте новое сообщение в веб-версии Outlook.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-114">In Outlook on the web, create a new message.</span></span>

1. <span data-ttu-id="2e6cf-115">Выберите \*\*... \*\* в нижней части нового сообщения, а затем выберите **Получить надстройки** в появившемся меню.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Окно создания сообщений в новой веб-версии Outlook с выделенной опцией "Получить надстройки"](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="2e6cf-117">В диалоговом окне **Надстройки для Outlook** выберите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Диалоговое окно "Надстройки для Outlook" в новой веб-версии Outlook с выбранной опцией "Мои надстройки "](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="2e6cf-119">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="2e6cf-120">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: управление надстройками с указанием параметра "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="2e6cf-p102">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="2e6cf-124">Загрузка неопубликованной надстройки в классической веб-версии Outlook</span><span class="sxs-lookup"><span data-stu-id="2e6cf-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="2e6cf-125">Перейдите к [Outlook в Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="2e6cf-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="2e6cf-126">Нажмите значок шестеренки в верхнем правом углу панели инструментов и выберите пункт **Управление надстройками**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Снимок экрана: веб-версия Outlook с параметром "Управление надстройками"](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="2e6cf-128">На странице **Управление надстройками** выберите **Надстройки** > **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Диалоговое окно магазина веб-версии Outlook с открытым разделом "Мои надстройки"](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="2e6cf-130">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="2e6cf-131">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: управление надстройками с указанием параметра "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="2e6cf-p104">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="2e6cf-135">Загрузка неопубликованной надстройки в классической версии Outlook</span><span class="sxs-lookup"><span data-stu-id="2e6cf-135">Sideload an add-in in Outlook on the desktop</span></span>

### <a name="outlook-2016-or-later"></a><span data-ttu-id="2e6cf-136">Outlook 2016 или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="2e6cf-136">Outlook 2016 or later</span></span>

1. <span data-ttu-id="2e6cf-137">Откройте Outlook 2016 или более поздней версии в Windows или Mac.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-137">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="2e6cf-138">Нажмите кнопку **Получить надстройки** на ленте.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-138">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Лента Outlook 2016 с указанием кнопки "Магазин"](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > <span data-ttu-id="2e6cf-140">Если вы не видите кнопку **Получить надстройки** в вашей версии Outlook, выберите кнопку **Store** на ленте.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-140">If you don't see the **Get Add-ins** button in your version of Outlook, select the **Store** button on the ribbon instead.</span></span>

1. <span data-ttu-id="2e6cf-141">Выберите **Надстройки**, а затем нажмите **Мои надстройки**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-141">Select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Диалоговое окно магазина Outlook 2016 с открытым разделом "Мои надстройки"](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="2e6cf-143">Найдите раздел **Пользовательские надстройки** в нижней части диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-143">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="2e6cf-144">Выберите **Добавить пользовательскую надстройку** > **Добавить из файла**.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Снимок экрана: магазин с параметром "Добавить из файла"](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="2e6cf-p106">Найдите файл манифеста для своей надстройки и установите его, подтверждая все запросы.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-2013"></a><span data-ttu-id="2e6cf-148">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="2e6cf-148">Outlook 2013</span></span>

1. <span data-ttu-id="2e6cf-149">Откройте Outlook 2013 в Windows.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-149">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="2e6cf-150">Выберите меню **файл** , а затем нажмите кнопку **Управление надстройками** на вкладке **сведения** . Outlook откроет браузер.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-150">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open a browser.</span></span>

1. <span data-ttu-id="2e6cf-151">Выполните действия, описанные в разделе [Загрузка неопубликованных надстройка в Outlook в Интернете,](#sideload-an-add-in-in-outlook-on-the-web) в соответствии с вашей версией Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-151">Follow the steps in the [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="2e6cf-152">Удаление надстройки неопубликованные</span><span class="sxs-lookup"><span data-stu-id="2e6cf-152">Remove a sideloaded add-in</span></span>

<span data-ttu-id="2e6cf-153">Чтобы удалить надстройку неопубликованные из Outlook, выполните действия, описанные в этой статье, чтобы найти надстройку в разделе **Настраиваемые** надстройки диалогового окна со списком установленных надстроек. Нажмите кнопку с многоточием ( `...` ) для надстройки, а затем нажмите кнопку **Удалить** , чтобы удалить эту надстройку.</span><span class="sxs-lookup"><span data-stu-id="2e6cf-153">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>