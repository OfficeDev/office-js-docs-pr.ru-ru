---
title: Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office
description: ''
ms.date: 09/09/2019
localization_priority: Priority
ms.openlocfilehash: 8c1a39e4574f7e8ea60cdf32ff3139d9b929fe5d
ms.sourcegitcommit: 24303ca235ebd7144a1d913511d8e4fb7c0e8c0d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2019
ms.locfileid: "36838531"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="96179-102">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="96179-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="96179-p101">Иногда при работе с вашими надстройками Office пользователи могут столкнуться с определенными проблемами. Например, надстройка может не загружаться или быть недоступной. Эта статья поможет вам устранить распространенные проблемы, с которыми сталкиваются пользователи при работе с вашими надстройками Office.</span><span class="sxs-lookup"><span data-stu-id="96179-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="96179-106">Для выявления и устранения проблем с надстройками также можно использовать [Fiddler](https://www.telerik.com/fiddler).</span><span class="sxs-lookup"><span data-stu-id="96179-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="96179-107">Распространенные ошибки и инструкции по устранению неполадок</span><span class="sxs-lookup"><span data-stu-id="96179-107">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="96179-108">В таблице ниже перечислены распространенные сообщения об ошибках, с которыми могут столкнуться пользователи, и действия, которые можно предпринять для устранения ошибки.</span><span class="sxs-lookup"><span data-stu-id="96179-108">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="96179-109">**Сообщение об ошибке**</span><span class="sxs-lookup"><span data-stu-id="96179-109">**Error message**</span></span>|<span data-ttu-id="96179-110">**Решение**</span><span class="sxs-lookup"><span data-stu-id="96179-110">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="96179-111">Ошибка приложения: не удалось подключиться к каталогу</span><span class="sxs-lookup"><span data-stu-id="96179-111">App error: Catalog could not be reached</span></span>|<span data-ttu-id="96179-p102">Проверьте настройки брандмауэра. Под каталогом понимается AppSource. Это сообщение означает, что пользователь не может получить доступ к AppSource.</span><span class="sxs-lookup"><span data-stu-id="96179-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="96179-p103">ОШИБКА ПРИЛОЖЕНИЯ: Нам не удалось запустить это приложение. Чтобы проигнорировать проблему, закройте данное окно. Чтобы попробовать еще раз, нажмите "Перезапустить".</span><span class="sxs-lookup"><span data-stu-id="96179-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="96179-116">Убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="96179-116">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="96179-117">Ошибка: объект не поддерживает свойство или метод 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="96179-117">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="96179-p104">Убедитесь, что Internet Explorer не работает в режиме совместимости. Откройте меню "Сервис" >  **Параметры просмотра в режиме совместимости**.</span><span class="sxs-lookup"><span data-stu-id="96179-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="96179-p105">К сожалению, нам не удалось загрузить приложение, так как ваша версия браузера не поддерживается. Щелкните здесь, чтобы открыть список поддерживаемых версий браузера.</span><span class="sxs-lookup"><span data-stu-id="96179-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="96179-p106">Убедитесь, что браузер поддерживает локальное хранилище HTML5, или сбросьте параметры Internet Explorer. Сведения о поддерживаемых браузерах см. в разделе [Требования к запуску надстроек для Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="96179-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="96179-124">При установке надстройки в строке состояния появляется сообщение "Ошибка при загрузке надстройки"</span><span class="sxs-lookup"><span data-stu-id="96179-124">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="96179-125">Закройте Office.</span><span class="sxs-lookup"><span data-stu-id="96179-125">Close Office.</span></span>
2. <span data-ttu-id="96179-126">Убедитесь, что манифест действителен.</span><span class="sxs-lookup"><span data-stu-id="96179-126">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="96179-127">Перезапустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="96179-127">Restart the add-in.</span></span>
4. <span data-ttu-id="96179-128">Переустановите надстройку.</span><span class="sxs-lookup"><span data-stu-id="96179-128">Install the add-in</span></span>

<span data-ttu-id="96179-129">Также можно отправить нам отзыв: при использовании Excel для Windows или Mac можно отправить отзыв группе расширяемости Office непосредственно из Excel.</span><span class="sxs-lookup"><span data-stu-id="96179-129">If using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="96179-130">Для этого выберите **Файл** | **Отзывы и предложения** | **Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="96179-130">To do this, select File -> Feedback -> Send a Frown.</span></span> <span data-ttu-id="96179-131">При отправке нахмуренного смайлика будут предоставлены необходимые журналы для понимания описываемой проблемы.</span><span class="sxs-lookup"><span data-stu-id="96179-131">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="96179-132">Надстройка Outlook работает неправильно</span><span class="sxs-lookup"><span data-stu-id="96179-132">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="96179-133">Если надстройка Outlook в Windows и [в Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) работает неправильно, попробуйте включить отладку сценариев в Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="96179-133">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="96179-134">Откройте меню "Сервис" >  **Свойства браузера** > **Дополнительно**.</span><span class="sxs-lookup"><span data-stu-id="96179-134">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="96179-135">В разделе  **Обзор**, снимите флажки  **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="96179-135">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="96179-p108">Снимать эти флажки рекомендуется только для устранения неполадки. Если они сняты, то при использовании браузера будут появляться соответствующие сообщения. После устранения проблемы снова установите флажки  **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="96179-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="96179-139">Надстройка не активируется в Office 2013</span><span class="sxs-lookup"><span data-stu-id="96179-139">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="96179-140">Если надстройка не активируется при выполнении пользователем следующих действий:</span><span class="sxs-lookup"><span data-stu-id="96179-140">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="96179-141">выполнении входа с помощью учетной записи Майкрософт в Office 2013;</span><span class="sxs-lookup"><span data-stu-id="96179-141">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="96179-142">включении двухшаговой проверки учетной записи Майкрософт;</span><span class="sxs-lookup"><span data-stu-id="96179-142">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="96179-143">проверки своего удостоверения по запросу при попытке добавления надстройки, —</span><span class="sxs-lookup"><span data-stu-id="96179-143">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="96179-144">убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="96179-144">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="96179-145">Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки</span><span class="sxs-lookup"><span data-stu-id="96179-145">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="96179-146">Сведения об устранении проблем, связанных с манифестом надстройки, см. в статье [Проверка манифеста и устранение связанных с ним неполадок](troubleshoot-manifest.md).</span><span class="sxs-lookup"><span data-stu-id="96179-146">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="96179-147">Не отображается диалоговое окно надстройки</span><span class="sxs-lookup"><span data-stu-id="96179-147">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="96179-p109">При открытии надстройки Office пользователю будет предложено разрешить отображение диалогового окна. Пользователь выбирает **Разрешить**, и появляется следующее сообщение об ошибке:</span><span class="sxs-lookup"><span data-stu-id="96179-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="96179-p110">"Параметры безопасности браузера не позволили создать диалоговое окно. Используйте другой браузер или настройте браузер так, чтобы [URL-адрес] и домен, отображаемый в адресной строке браузера, находились в одной зоне безопасности."</span><span class="sxs-lookup"><span data-stu-id="96179-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Снимок экрана: сообщение об ошибке](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="96179-153">**Браузеры**</span><span class="sxs-lookup"><span data-stu-id="96179-153">**Affected browsers**</span></span>|<span data-ttu-id="96179-154">**Платформы**</span><span class="sxs-lookup"><span data-stu-id="96179-154">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="96179-155">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="96179-155">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="96179-156">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="96179-156">Office on the web</span></span>|

<span data-ttu-id="96179-p111">Чтобы решить эту проблему, пользователи или администраторы могут добавить домен надстройки в список надежных сайтов в Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="96179-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="96179-159">Не добавляйте URL-адрес надстройки в список надежных сайтов, если вы не доверяете надстройке.</span><span class="sxs-lookup"><span data-stu-id="96179-159">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="96179-160">Чтобы добавить URL-адрес в список надежных сайтов:</span><span class="sxs-lookup"><span data-stu-id="96179-160">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="96179-161">На **панели управления** перейдите в раздел **Свойства браузера** > **Безопасность**.</span><span class="sxs-lookup"><span data-stu-id="96179-161">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="96179-162">Выберите зону **Надежные сайты** и нажмите кнопку **Сайты**.</span><span class="sxs-lookup"><span data-stu-id="96179-162">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="96179-163">Введите URL-адрес из сообщения об ошибке и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="96179-163">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="96179-p112">Запустите надстройку снова. Если проблема не исчезла, проверьте параметры для других зон безопасности и убедитесь, что домен надстройки находится в той же зоне, что и URL-адрес, отображаемый в адресной строке приложения Office.</span><span class="sxs-lookup"><span data-stu-id="96179-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="96179-p113">Эта проблема возникает при использовании Dialog API в режиме всплывающих окон. Чтобы эта проблема не возникала, используйте флажок [displayInFrame](/javascript/api/office/office.ui). Для этого страница должна поддерживать отображение в окнах iframe. В приведенном ниже примере показано, как использовать флажок.</span><span class="sxs-lookup"><span data-stu-id="96179-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="96179-170">Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются</span><span class="sxs-lookup"><span data-stu-id="96179-170">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="96179-171">Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере.</span><span class="sxs-lookup"><span data-stu-id="96179-171">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="96179-172">Для Windows:</span><span class="sxs-lookup"><span data-stu-id="96179-172">For Windows:</span></span>
<span data-ttu-id="96179-173">Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="96179-173">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="96179-174">Для Mac</span><span class="sxs-lookup"><span data-stu-id="96179-174">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="96179-175">Для iOS</span><span class="sxs-lookup"><span data-stu-id="96179-175">For iOS:</span></span>
<span data-ttu-id="96179-p114">Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="96179-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="96179-178">См. также</span><span class="sxs-lookup"><span data-stu-id="96179-178">See also</span></span>

- [<span data-ttu-id="96179-179">Отладка надстроек в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="96179-179">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="96179-180">Загрузка неопубликованной надстройки Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="96179-180">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="96179-181">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="96179-181">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="96179-182">Проверка манифеста и устранение связанных с ним неполадок</span><span class="sxs-lookup"><span data-stu-id="96179-182">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
