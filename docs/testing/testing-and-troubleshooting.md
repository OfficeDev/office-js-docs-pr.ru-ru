---
title: Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office
description: Узнайте, как устранять ошибки пользователей в надстройках Office.
ms.date: 04/13/2020
localization_priority: Normal
ms.openlocfilehash: 81ca9eac89fb949bd2c1f7f63861d8a4f5295a79
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241051"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="fd318-103">Устранение ошибок, с которыми сталкиваются пользователи при работе с надстройками Office</span><span class="sxs-lookup"><span data-stu-id="fd318-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="fd318-p101">Иногда при работе с вашими надстройками Office пользователи могут столкнуться с определенными проблемами. Например, надстройка может не загружаться или быть недоступной. Эта статья поможет вам устранить распространенные проблемы, с которыми сталкиваются пользователи при работе с вашими надстройками Office.</span><span class="sxs-lookup"><span data-stu-id="fd318-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="fd318-107">Для выявления и устранения проблем с надстройками также можно использовать [Fiddler](https://www.telerik.com/fiddler).</span><span class="sxs-lookup"><span data-stu-id="fd318-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="fd318-108">Распространенные ошибки и инструкции по устранению неполадок</span><span class="sxs-lookup"><span data-stu-id="fd318-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="fd318-109">В таблице ниже перечислены распространенные сообщения об ошибках, с которыми могут столкнуться пользователи, и действия, которые можно предпринять для устранения ошибки.</span><span class="sxs-lookup"><span data-stu-id="fd318-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="fd318-110">**Сообщение об ошибке**</span><span class="sxs-lookup"><span data-stu-id="fd318-110">**Error message**</span></span>|<span data-ttu-id="fd318-111">**Решение**</span><span class="sxs-lookup"><span data-stu-id="fd318-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="fd318-112">Ошибка приложения: не удалось подключиться к каталогу</span><span class="sxs-lookup"><span data-stu-id="fd318-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="fd318-p102">Проверьте настройки брандмауэра. Под каталогом понимается AppSource. Это сообщение означает, что пользователь не может получить доступ к AppSource.</span><span class="sxs-lookup"><span data-stu-id="fd318-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="fd318-p103">ОШИБКА ПРИЛОЖЕНИЯ: Нам не удалось запустить это приложение. Чтобы проигнорировать проблему, закройте данное окно. Чтобы попробовать еще раз, нажмите "Перезапустить".</span><span class="sxs-lookup"><span data-stu-id="fd318-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="fd318-117">Убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="fd318-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="fd318-118">Ошибка: объект не поддерживает свойство или метод 'defineProperty'</span><span class="sxs-lookup"><span data-stu-id="fd318-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="fd318-119">Убедитесь, что Internet Explorer не работает в режиме совместимости.</span><span class="sxs-lookup"><span data-stu-id="fd318-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="fd318-120">Откройте меню "Сервис" > **Параметры просмотра в режиме совместимости**.</span><span class="sxs-lookup"><span data-stu-id="fd318-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="fd318-p105">К сожалению, нам не удалось загрузить приложение, так как ваша версия браузера не поддерживается. Щелкните здесь, чтобы открыть список поддерживаемых версий браузера.</span><span class="sxs-lookup"><span data-stu-id="fd318-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="fd318-p106">Убедитесь, что браузер поддерживает локальное хранилище HTML5, или сбросьте параметры Internet Explorer. Сведения о поддерживаемых браузерах см. в разделе [Требования к запуску надстроек для Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="fd318-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="fd318-125">При установке надстройки в строке состояния появляется сообщение "Ошибка при загрузке надстройки"</span><span class="sxs-lookup"><span data-stu-id="fd318-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="fd318-126">Закройте Office.</span><span class="sxs-lookup"><span data-stu-id="fd318-126">Close Office.</span></span>
2. <span data-ttu-id="fd318-127">Убедитесь, что манифест действителен.</span><span class="sxs-lookup"><span data-stu-id="fd318-127">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="fd318-128">Перезапустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="fd318-128">Restart the add-in</span></span>
4. <span data-ttu-id="fd318-129">Переустановите надстройку.</span><span class="sxs-lookup"><span data-stu-id="fd318-129">Install the add-in again.</span></span>

<span data-ttu-id="fd318-130">Также можно отправить нам отзыв: при использовании Excel для Windows или Mac можно отправить отзыв группе расширяемости Office непосредственно из Excel.</span><span class="sxs-lookup"><span data-stu-id="fd318-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="fd318-131">Для этого выберите **Файл** | **Отзывы и предложения** | **Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="fd318-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="fd318-132">При отправке нахмуренного смайлика будут предоставлены необходимые журналы для понимания описываемой проблемы.</span><span class="sxs-lookup"><span data-stu-id="fd318-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="fd318-133">Надстройка Outlook работает неправильно</span><span class="sxs-lookup"><span data-stu-id="fd318-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="fd318-134">Если надстройка Outlook в Windows и [в Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) работает неправильно, попробуйте включить отладку сценариев в Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="fd318-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="fd318-135">Откройте раздел Сервис > " **Свойства** > браузера **".**</span><span class="sxs-lookup"><span data-stu-id="fd318-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="fd318-136">В разделе **Обзор**, снимите флажки **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="fd318-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="fd318-p108">Снимать эти флажки рекомендуется только для устранения неполадки. Если они сняты, то при использовании браузера будут появляться соответствующие сообщения. После устранения проблемы снова установите флажки **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="fd318-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="fd318-140">Надстройка не активируется в Office 2013</span><span class="sxs-lookup"><span data-stu-id="fd318-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="fd318-141">Если надстройка не активируется при выполнении пользователем следующих действий:</span><span class="sxs-lookup"><span data-stu-id="fd318-141">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="fd318-142">выполнении входа с помощью учетной записи Майкрософт в Office 2013;</span><span class="sxs-lookup"><span data-stu-id="fd318-142">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="fd318-143">включении двухшаговой проверки учетной записи Майкрософт;</span><span class="sxs-lookup"><span data-stu-id="fd318-143">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="fd318-144">проверки своего удостоверения по запросу при попытке добавления надстройки, —</span><span class="sxs-lookup"><span data-stu-id="fd318-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="fd318-145">убедитесь, что установлены последние обновления Office, или скачайте [обновление для Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="fd318-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="fd318-146">Надстройка не загружается в область задач или возникают другие проблемы с манифестом надстройки</span><span class="sxs-lookup"><span data-stu-id="fd318-146">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="fd318-147">Сведения об отладке проблем с манифестом см. в статьях [Проверка манифеста надстройки Office](troubleshoot-manifest.md) и [Отладка надстройки с помощью журнала среды выполнения](runtime-logging.md).</span><span class="sxs-lookup"><span data-stu-id="fd318-147">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="fd318-148">Не отображается диалоговое окно надстройки</span><span class="sxs-lookup"><span data-stu-id="fd318-148">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="fd318-p109">При открытии надстройки Office пользователю будет предложено разрешить отображение диалогового окна. Пользователь выбирает **Разрешить**, и появляется следующее сообщение об ошибке:</span><span class="sxs-lookup"><span data-stu-id="fd318-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="fd318-p110">"Параметры безопасности браузера не позволили создать диалоговое окно. Используйте другой браузер или настройте браузер так, чтобы [URL-адрес] и домен, отображаемый в адресной строке браузера, находились в одной зоне безопасности."</span><span class="sxs-lookup"><span data-stu-id="fd318-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Снимок экрана: сообщение об ошибке](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="fd318-154">**Браузеры**</span><span class="sxs-lookup"><span data-stu-id="fd318-154">**Affected browsers**</span></span>|<span data-ttu-id="fd318-155">**Платформы**</span><span class="sxs-lookup"><span data-stu-id="fd318-155">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="fd318-156">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="fd318-156">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="fd318-157">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fd318-157">Office on the web</span></span>|

<span data-ttu-id="fd318-p111">Чтобы решить эту проблему, пользователи или администраторы могут добавить домен надстройки в список надежных сайтов в Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="fd318-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="fd318-160">Не добавляйте URL-адрес надстройки в список надежных сайтов, если вы не доверяете надстройке.</span><span class="sxs-lookup"><span data-stu-id="fd318-160">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="fd318-161">Чтобы добавить URL-адрес в список надежных сайтов:</span><span class="sxs-lookup"><span data-stu-id="fd318-161">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="fd318-162">На **панели управления** перейдите в раздел **Свойства браузера** > **Безопасность**.</span><span class="sxs-lookup"><span data-stu-id="fd318-162">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="fd318-163">Выберите зону **Надежные сайты** и нажмите кнопку **Сайты**.</span><span class="sxs-lookup"><span data-stu-id="fd318-163">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="fd318-164">Введите URL-адрес из сообщения об ошибке и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="fd318-164">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="fd318-p112">Запустите надстройку снова. Если проблема не исчезла, проверьте параметры для других зон безопасности и убедитесь, что домен надстройки находится в той же зоне, что и URL-адрес, отображаемый в адресной строке приложения Office.</span><span class="sxs-lookup"><span data-stu-id="fd318-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="fd318-p113">Эта проблема возникает при использовании Dialog API в режиме всплывающих окон. Чтобы эта проблема не возникала, используйте флажок [displayInFrame](/javascript/api/office/office.ui). Для этого страница должна поддерживать отображение в окнах iframe. В приведенном ниже примере показано, как использовать флажок.</span><span class="sxs-lookup"><span data-stu-id="fd318-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="fd318-171">Изменения команд надстройки, в том числе кнопок ленты и элементов меню, не отображаются</span><span class="sxs-lookup"><span data-stu-id="fd318-171">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="fd318-172">Если внесенные в манифест изменения (например, имена значков кнопок на ленте или текст элементов меню) не вступили в силу, попробуйте очистить кэш Office на своем компьютере.</span><span class="sxs-lookup"><span data-stu-id="fd318-172">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="fd318-173">Для Windows:</span><span class="sxs-lookup"><span data-stu-id="fd318-173">For Windows:</span></span>
<span data-ttu-id="fd318-174">Удалите содержимое папки `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="fd318-174">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="fd318-175">Для Mac</span><span class="sxs-lookup"><span data-stu-id="fd318-175">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="fd318-176">Для iOS</span><span class="sxs-lookup"><span data-stu-id="fd318-176">For iOS:</span></span>
<span data-ttu-id="fd318-p114">Для принудительной перезагрузки вызовите метод JavaScript `window.location.reload(true)` в надстройке. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="fd318-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="fd318-179">Изменения статических файлов, таких как JavaScript, HTML и CSS, не отображаются.</span><span class="sxs-lookup"><span data-stu-id="fd318-179">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="fd318-180">Браузер может кэшировать эти файлы.</span><span class="sxs-lookup"><span data-stu-id="fd318-180">The browser may be caching these files.</span></span> <span data-ttu-id="fd318-181">Чтобы избежать этого, отключите кэширование на стороне клиента при разработке.</span><span class="sxs-lookup"><span data-stu-id="fd318-181">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="fd318-182">Сведения будут зависеть от того, какой тип сервера вы используете.</span><span class="sxs-lookup"><span data-stu-id="fd318-182">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="fd318-183">В большинстве случаев необходимо добавить определенные заголовки в HTTP-ответы.</span><span class="sxs-lookup"><span data-stu-id="fd318-183">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="fd318-184">Мы предлагаем следующий набор заголовков:</span><span class="sxs-lookup"><span data-stu-id="fd318-184">We suggest the following set:</span></span>

- <span data-ttu-id="fd318-185">Cache-Control: "private, no-cache, no-store"</span><span class="sxs-lookup"><span data-stu-id="fd318-185">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="fd318-186">Pragma: "No-cache"</span><span class="sxs-lookup"><span data-stu-id="fd318-186">Pragma: "no-cache"</span></span>
- <span data-ttu-id="fd318-187">Expires: "-1"</span><span class="sxs-lookup"><span data-stu-id="fd318-187">Expires: "-1"</span></span>

<span data-ttu-id="fd318-188">Пример использования на сервере Node.JS Express см. в [этом файле app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="fd318-188">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="fd318-189">Пример использования в проекте ASP.NET см. в [этом файле cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="fd318-189">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="fd318-190">Если надстройка размещена на сервере Internet Information Server (IIS), можно также добавить указанные сведения в файл web.config.</span><span class="sxs-lookup"><span data-stu-id="fd318-190">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="fd318-191">Если сначала эти действия безуспешны, вам, возможно, потребуется очистить кэш браузера.</span><span class="sxs-lookup"><span data-stu-id="fd318-191">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="fd318-192">Сделайте это в интерфейсе браузера.</span><span class="sxs-lookup"><span data-stu-id="fd318-192">Do this through the UI of the browser.</span></span> <span data-ttu-id="fd318-193">Иногда очистить кэш браузера Microsoft Edge, используя пользовательский интерфейс, не удается.</span><span class="sxs-lookup"><span data-stu-id="fd318-193">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="fd318-194">В таком случае выполните следующую команду в командной строке Windows.</span><span class="sxs-lookup"><span data-stu-id="fd318-194">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a><span data-ttu-id="fd318-195">См. также</span><span class="sxs-lookup"><span data-stu-id="fd318-195">See also</span></span>

- [<span data-ttu-id="fd318-196">Отладка надстроек в Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="fd318-196">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="fd318-197">Загрузка неопубликованной надстройки Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="fd318-197">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="fd318-198">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="fd318-198">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="fd318-199">Проверка манифеста надстройки Office</span><span class="sxs-lookup"><span data-stu-id="fd318-199">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="fd318-200">Отладка надстройки с помощью журнала среды выполнения</span><span class="sxs-lookup"><span data-stu-id="fd318-200">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
