---
title: Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: 375b3819d423362c7d5e124700a0bea2dcf6e9e0
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19438825"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="328e4-102">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="328e4-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="328e4-p101">Иногда при работе с вашими надстройками Office пользователи могут столкнуться с определенными проблемами. Например, надстройка может не загружаться или быть недоступной. Эта статья поможет вам устранить распространенные проблемы, с которыми сталкиваются пользователи при работе с вашими надстройками Office.</span><span class="sxs-lookup"><span data-stu-id="328e4-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="328e4-106">Для выявления и устранения проблем с надстройками также можно использовать [Fiddler](http://www.telerik.com/fiddler).</span><span class="sxs-lookup"><span data-stu-id="328e4-106">You can also use [Fiddler](http://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="328e4-107">Устранив проблему, вы можете [написать об этом пользователям в AppSource напрямую](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="328e4-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](https://docs.microsoft.com/en-us/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="328e4-108">Распространенные ошибки и инструкции по устранению неполадок</span><span class="sxs-lookup"><span data-stu-id="328e4-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="328e4-109">В таблице ниже перечислены распространенные сообщения об ошибках, с которыми могут столкнуться пользователи, и действия, которые можно предпринять для устранения ошибки.</span><span class="sxs-lookup"><span data-stu-id="328e4-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="328e4-110">**Сообщение об ошибке**</span><span class="sxs-lookup"><span data-stu-id="328e4-110">**Error message**</span></span>|<span data-ttu-id="328e4-111">**Решение**</span><span class="sxs-lookup"><span data-stu-id="328e4-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="328e4-112">Ошибка приложения: не удалось подключиться к каталогу</span><span class="sxs-lookup"><span data-stu-id="328e4-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="328e4-p102">Проверьте настройки брандмауэра. Под каталогом понимается AppSource. Это сообщение означает, что пользователь не может получить доступ к AppSource.</span><span class="sxs-lookup"><span data-stu-id="328e4-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="328e4-p103">ОШИБКА ПРИЛОЖЕНИЯ: Нам не удалось запустить это приложение. Чтобы проигнорировать проблему, закройте данное окно. Чтобы попробовать еще раз, нажмите "Перезапустить".</span><span class="sxs-lookup"><span data-stu-id="328e4-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="328e4-117">Убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="328e4-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span></span>|
|<span data-ttu-id="328e4-118">Ошибка: объект не поддерживает свойство или метод 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="328e4-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="328e4-p104">Убедитесь, что Internet Explorer не работает в режиме совместимости. Откройте меню "Сервис" >  **Параметры просмотра в режиме совместимости**.</span><span class="sxs-lookup"><span data-stu-id="328e4-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="328e4-p105">К сожалению, нам не удалось загрузить приложение, так как ваша версия браузера не поддерживается. Щелкните здесь, чтобы открыть список поддерживаемых версий браузера.</span><span class="sxs-lookup"><span data-stu-id="328e4-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="328e4-p106">Убедитесь, что браузер поддерживает локальное хранилище HTML5, или сбросьте параметры Internet Explorer. Сведения о поддерживаемых браузерах см. в разделе [Требования к запуску надстроек для Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="328e4-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="328e4-125">Надстройка Outlook работает неправильно</span><span class="sxs-lookup"><span data-stu-id="328e4-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="328e4-126">Если надстройка Outlook в Windows работает неправильно, попробуйте включить отладку сценариев в Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="328e4-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="328e4-127">Откройте меню "Сервис" >  **Свойства браузера** > **Дополнительно**.</span><span class="sxs-lookup"><span data-stu-id="328e4-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="328e4-128">В разделе  **Обзор**, снимите флажки  **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="328e4-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="328e4-p107">Снимать эти флажки рекомендуется только для устранения неполадки. Если они сняты, то при использовании браузера будут появляться соответствующие сообщения. После устранения проблемы снова установите флажки  **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="328e4-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="328e4-132">Надстройка не активируется в Office 2013</span><span class="sxs-lookup"><span data-stu-id="328e4-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="328e4-133">Если надстройка не активируется при выполнении пользователем следующих действий:</span><span class="sxs-lookup"><span data-stu-id="328e4-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="328e4-134">выполнении входа с помощью учетной записи Майкрософт в Office 2013;</span><span class="sxs-lookup"><span data-stu-id="328e4-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="328e4-135">включении двухшаговой проверки учетной записи Майкрософт;</span><span class="sxs-lookup"><span data-stu-id="328e4-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="328e4-136">проверки своего удостоверения по запросу при попытке добавления надстройки, —</span><span class="sxs-lookup"><span data-stu-id="328e4-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="328e4-137">убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="328e4-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/en-us/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="328e4-138">Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки</span><span class="sxs-lookup"><span data-stu-id="328e4-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="328e4-139">Сведения об устранении проблем, связанных с манифестом надстройки, см. в статье [Проверка манифеста и устранение связанных с ним неполадок](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="328e4-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="328e4-140">Не отображается диалоговое окно надстройки</span><span class="sxs-lookup"><span data-stu-id="328e4-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="328e4-p108">При открытии надстройки Office пользователю будет предложено разрешить отображение диалогового окна. Пользователь выбирает **Разрешить**, и появляется следующее сообщение об ошибке:</span><span class="sxs-lookup"><span data-stu-id="328e4-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="328e4-p109">"Параметры безопасности браузера не позволили создать диалоговое окно. Используйте другой браузер или настройте браузер так, чтобы [URL-адрес] и домен, отображаемый в адресной строке браузера, находились в одной зоне безопасности."</span><span class="sxs-lookup"><span data-stu-id="328e4-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Снимок экрана: сообщение об ошибке](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="328e4-146">**Браузеры**</span><span class="sxs-lookup"><span data-stu-id="328e4-146">**Affected browsers**</span></span>|<span data-ttu-id="328e4-147">**Платформы**</span><span class="sxs-lookup"><span data-stu-id="328e4-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="328e4-148">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="328e4-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="328e4-149">Office Online</span><span class="sxs-lookup"><span data-stu-id="328e4-149">Office Online</span></span>|

<span data-ttu-id="328e4-p110">Чтобы решить эту проблему, пользователи или администраторы могут добавить домен надстройки в список надежных сайтов в Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="328e4-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="328e4-152">Не добавляйте URL-адрес надстройки в список надежных сайтов, если вы не доверяете надстройке.</span><span class="sxs-lookup"><span data-stu-id="328e4-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="328e4-153">Чтобы добавить URL-адрес в список надежных сайтов:</span><span class="sxs-lookup"><span data-stu-id="328e4-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="328e4-154">В Internet Explorer нажмите кнопку "Сервис" и перейдите в раздел **Свойства браузера** > **Безопасность**.</span><span class="sxs-lookup"><span data-stu-id="328e4-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="328e4-155">Выберите зону **Надежные сайты** и нажмите кнопку **Сайты**.</span><span class="sxs-lookup"><span data-stu-id="328e4-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="328e4-156">Введите URL-адрес из сообщения об ошибке и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="328e4-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="328e4-p111">Запустите надстройку снова. Если проблема не исчезла, проверьте параметры для других зон безопасности и убедитесь, что домен надстройки находится в той же зоне, что и URL-адрес, отображаемый в адресной строке приложения Office.</span><span class="sxs-lookup"><span data-stu-id="328e4-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="328e4-p112">Эта проблема возникает при использовании Dialog API в режиме всплывающих окон. Чтобы эта проблема не возникала, используйте флажок [displayInFrame](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync). Для этого страница должна поддерживать отображение в окнах iframe. В приведенном ниже примере показано, как использовать флажок.</span><span class="sxs-lookup"><span data-stu-id="328e4-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](https://dev.office.com/reference/add-ins/shared/officeui.displaydialogasync) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js

Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="328e4-163">Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются</span><span class="sxs-lookup"><span data-stu-id="328e4-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>
<span data-ttu-id="328e4-p113">Иногда изменения таких команд надстройки, как значок кнопки на ленте или текст элемента меню, не отображаются. Удалите из кэша Office старые версии.</span><span class="sxs-lookup"><span data-stu-id="328e4-p113">Sometimes changes to add-in commands such as the icon for a ribbon button or the text of a menu item do not seem to take effect. Clear the Office cache of the old versions.</span></span>

#### <a name="for-windows"></a><span data-ttu-id="328e4-166">Для Windows:</span><span class="sxs-lookup"><span data-stu-id="328e4-166">For Windows:</span></span>
<span data-ttu-id="328e4-167">Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="328e4-167">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="328e4-168">Для Mac</span><span class="sxs-lookup"><span data-stu-id="328e4-168">For Mac:</span></span>
<span data-ttu-id="328e4-169">Удалите содержимое папки `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="328e4-169">Delete the content of the folder `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span>

#### <a name="for-ios"></a><span data-ttu-id="328e4-170">Для iOS</span><span class="sxs-lookup"><span data-stu-id="328e4-170">For iOS:</span></span>
<span data-ttu-id="328e4-p114">Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="328e4-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="328e4-173">См. также</span><span class="sxs-lookup"><span data-stu-id="328e4-173">See also</span></span>

- [<span data-ttu-id="328e4-174">Отладка надстроек в Office Online</span><span class="sxs-lookup"><span data-stu-id="328e4-174">Debug add-ins in Office Online</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="328e4-175">Загрузка неопубликованной надстройки Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="328e4-175">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="328e4-176">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="328e4-176">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="328e4-177">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="328e4-177">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
