---
title: Отладка надстроек Office на iPad и Mac
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: 587aa48c65682a4ccaedadfcc550c84850b87976
ms.sourcegitcommit: 9b021af6cb23a58486d6c5c7492be425e309bea1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/15/2018
ms.locfileid: "26533744"
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a><span data-ttu-id="e89cf-102">Отладка надстроек Office на iPad и Mac</span><span class="sxs-lookup"><span data-stu-id="e89cf-102">Debug Office Add-ins on iPad and Mac</span></span>

<span data-ttu-id="e89cf-p101">Visual Studio подходит для разработки и отладки надстроек в Windows, но с его помощью невозможно выполнять отладку надстроек на платформах iPad и Mac. Так как надстройки создаются с помощью кода HTML и Javascript, они рассчитаны на работу на многих платформах, но отрисовка HTML в разных браузерах может слегка отличаться. В этой статье описывается отладка надстроек на платформах iPad и Mac.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p101">You can use Visual Studio to develop and debug add-ins on Windows, but you can't use it to debug add-ins on the iPad or Mac. Because add-ins are developed using HTML and Javascript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on an iPad or Mac.</span></span> 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="e89cf-106">Отладка с помощью Safari Web Inspector на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="e89cf-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="e89cf-107">Если у вас есть надстройка, которая отображает пользовательский интерфейс в области задач или контентной надстройке, вы можете отлаживать надстройку Office с помощью Safari Web Inspector.</span><span class="sxs-lookup"><span data-stu-id="e89cf-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="e89cf-108">Отладку надстроек Office на компьютере Mac можно выполнить, только если на нем установлена система Mac OS High Sierra И Office для Mac версии 16.9.1 (сборка 18012504) или более поздней.</span><span class="sxs-lookup"><span data-stu-id="e89cf-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="e89cf-109">Если у вас нет сборки Office для Mac, вы можете получить ее, присоединившись к [программе для разработчиков Office 365](https://aka.ms/o365devprogram).</span><span class="sxs-lookup"><span data-stu-id="e89cf-109">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://aka.ms/o365devprogram).</span></span>

<span data-ttu-id="e89cf-110">Для этого откройте терминал и установите свойство `OfficeWebAddinDeveloperExtras` для соответствующего приложения Office следующим образом:</span><span class="sxs-lookup"><span data-stu-id="e89cf-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="e89cf-111">Затем откройте приложение Office и вставьте надстройку.</span><span class="sxs-lookup"><span data-stu-id="e89cf-111">Then, open the Office application and insert your add-in.</span></span> <span data-ttu-id="e89cf-112">Щелкните надстройку правой кнопкой мыши. В контекстном меню отобразится пункт **Проверить элемент**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-112">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span>  <span data-ttu-id="e89cf-113">Выберите его. Он появится в инспекторе, где можно устанавливать точки останова и отлаживать надстройку.</span><span class="sxs-lookup"><span data-stu-id="e89cf-113">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e89cf-114">Обратите внимание, что это экспериментальная функция, которая может отсутствовать в последующих версиях приложений Office.</span><span class="sxs-lookup"><span data-stu-id="e89cf-114">Please note that this is an experimental feature and there are no guarantees that we will preserve this functionality in future versions of Office applications.</span></span>
>
> <span data-ttu-id="e89cf-115">Если при попытке использовать инспектор диалоговое окно мерцает, попробуйте сделать вот что:</span><span class="sxs-lookup"><span data-stu-id="e89cf-115">If you are trying to use the inspector and the dialog flickers, try the following workaround:</span></span>
> 1. <span data-ttu-id="e89cf-116">Уменьшите размер диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="e89cf-116">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="e89cf-117">Выберите пункт **Проверить элемент**, который откроется в новом окне.</span><span class="sxs-lookup"><span data-stu-id="e89cf-117">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="e89cf-118">Измените размер диалогового окна на исходный.</span><span class="sxs-lookup"><span data-stu-id="e89cf-118">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="e89cf-119">Используйте инспектор должным образом.</span><span class="sxs-lookup"><span data-stu-id="e89cf-119">Use the inspector as required.</span></span>

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a><span data-ttu-id="e89cf-120">Отладка с помощью Vorlon.JS на iPad или Mac</span><span class="sxs-lookup"><span data-stu-id="e89cf-120">Debugging with Vorlon.JS on a iPad or Mac</span></span>

<span data-ttu-id="e89cf-121">Чтобы отладить надстройку на iPad или Mac, можно использовать отладчик для веб-страниц Vorlon.JS, который схож с инструментами F12.</span><span class="sxs-lookup"><span data-stu-id="e89cf-121">To debug an add-in on iPad or Mac, you can use Vorlon.JS, a debugger for web pages that is similar to the F12 tools.</span></span> <span data-ttu-id="e89cf-122">Он предназначен для удаленной работы и позволяет отлаживать веб-страницы на различных устройствах.</span><span class="sxs-lookup"><span data-stu-id="e89cf-122">It is designed to work remotely and it enables you to debug web pages across different devices.</span></span> <span data-ttu-id="e89cf-123">Дополнительные сведения см. на [веб-сайте Vorlon](http://www.vorlonjs.com).</span><span class="sxs-lookup"><span data-stu-id="e89cf-123">For more information, see the [Vorlon website](http://www.vorlonjs.com).</span></span>  


### <a name="install-and-set-up-vorlonjs"></a><span data-ttu-id="e89cf-124">Установка и настройка Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="e89cf-124">Install and set up up Vorlon.JS on a Mac or iPad</span></span>  

1.  <span data-ttu-id="e89cf-125">Войдите на устройстве от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="e89cf-125">Log on to the device as an administrator.</span></span>

2.  <span data-ttu-id="e89cf-126">Установите библиотеку [Node.js](https://nodejs.org), если она еще не установлена.</span><span class="sxs-lookup"><span data-stu-id="e89cf-126">Install [Node.js](https://nodejs.org) if it isn't already installed.</span></span>

3.  <span data-ttu-id="e89cf-p105">Откройте окно **Терминал** и введите команду `npm i -g vorlon`. Средство будет установлено в папке `/usr/local/lib/node_modules/vorlon`.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p105">Open a **Terminal** window and enter the command `npm i -g vorlon`. The tool is installed to `/usr/local/lib/node_modules/vorlon`.</span></span>


### <a name="configure-vorlonjs-to-use-https"></a><span data-ttu-id="e89cf-129">Настройка Vorlon.JS на использование HTTPS</span><span class="sxs-lookup"><span data-stu-id="e89cf-129">Configure Vorlon.JS to use HTTPS</span></span>

<span data-ttu-id="e89cf-p106">Для отладки приложения с помощью Vorlon.JS необходимо добавить на начальную страницу приложения тег `<script>`, который загружает скрипт Vorlon.JS из известного расположения (подробные сведения см. в следующей процедуре). Если для надстройки используется протокол SSL (HTTPS), все используемые ею скрипты, включая скрипт Vorlon.JS, должны размещаться на HTTPS-сервере. Таким образом, чтобы использовать Vorlon.JS с надстройками, необходимо настроить Vorlon.JS на применение SSL.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p106">To debug an application using Vorlon.JS, you add a `<script>` tag to the opening page of the application that loads a Vorlon.JS script from a well-known location (for details, see the following procedure). If an add-in is SSL-secured (HTTPS), any scripts that it uses must be hosted from an HTTPS server, including the Vorlon.JS script. Therefore, you must configure Vorlon.JS to use SSL in order to use Vorlon.JS with add-ins.</span></span>

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  <span data-ttu-id="e89cf-133">В средстве **Finder** перейдите к папке `/usr/local/lib/node_modules/vorlon`, откройте контекстное меню (правая кнопка мыши) папки `/Server` и выберите пункт **Получить сведения**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-133">In **Finder**, go to `/usr/local/lib/node_modules/vorlon`, open the context menu for (right-click) the `/Server` folder, and then select **Get Info**.</span></span>

2.  <span data-ttu-id="e89cf-134">Выберите значок висячего замка в правом нижнем углу окна **Сведения о сервере**, чтобы разблокировать папку.</span><span class="sxs-lookup"><span data-stu-id="e89cf-134">Choose the padlock icon in the lower right corner of the **Server info** window to unlock the folder.</span></span>

3. <span data-ttu-id="e89cf-135">В разделе **Общий доступ и разрешения** этого окна выберите **Чтение и запись** в разделе **Разрешение** для группы **Персонал**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-135">In the **Sharing and Permissions** section of the window, set the **Privilege** for the **staff** group to **Read & Write**.</span></span>

4. <span data-ttu-id="e89cf-136">Снова выберите значок висячего замка, чтобы ***повторно заблокировать*** папку.</span><span class="sxs-lookup"><span data-stu-id="e89cf-136">Choose the padlock icon again to ***relock*** the folder.</span></span>

5. <span data-ttu-id="e89cf-137">В средстве **Finder** разверните вложенную папку `/Server`, щелкните правой кнопкой мыши файл `config.json` и выберите пункт **Получить сведения**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-137">Back in **Finder**, expand the `/Server` subfolder, right-click the file `config.json`, and then select **Get Info**.</span></span>

6. <span data-ttu-id="e89cf-p107">В окне **Сведения о config.json** измените разрешения для файла точно так же, как и для его родительской папки `/Server`. Не забудьте повторно заблокировать папку, а затем закройте окно.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p107">In the **config.json info** window, change the privileges of the file exactly the way you did for its parent `/Server` folder. Be sure to relock and close the window.</span></span>

7. <span data-ttu-id="e89cf-p108">В средстве **Finder** щелкните правой кнопкой мыши файл `config.json`, выберите пункт **Открыть с помощью** и выберите **TextEdit**. Файл откроется в текстовом редакторе.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p108">Back in **Finder**, right-click the file `config.json`, select **Open with**, and then select **TextEdit**. The file opens in a text editor.</span></span>

8. <span data-ttu-id="e89cf-142">Измените значение свойства **useSSL** на `true`.</span><span class="sxs-lookup"><span data-stu-id="e89cf-142">Change the value of the **useSSL** property to `true`.</span></span>

9. <span data-ttu-id="e89cf-p109">В разделе **Подключаемые модули** найдите подключаемый модуль, свойство **id** которого равно `OFFICE`, а свойство **name** — `Office Addin`. Если для свойства **enabled** этого подключаемого модуля еще не задано значение `true`, задайте для него значение `true`.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p109">In the **plugins** section, find the plugin with the **id** of `OFFICE` and the **name** of `Office Addin`. If the **enabled** property for the plug-in is not already `true`, set it to `true`.</span></span>

10. <span data-ttu-id="e89cf-145">Сохраните файл и закройте редактор.</span><span class="sxs-lookup"><span data-stu-id="e89cf-145">Save the file and close the editor.</span></span>

11. <span data-ttu-id="e89cf-146">В средстве **Finder** перейдите к папке `/usr/local/lib/node_modules/vorlon`, щелкните правой кнопкой мыши вложенную папку `Server` и выберите пункт **Новый терминал в этой папке**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-146">In **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span>

12. <span data-ttu-id="e89cf-p110">В окне **Терминал** введите команду `sudo vorlon`. Вам будет предложено ввести пароль администратора. Запустится сервер Vorlon. Оставьте окно **Терминал** открытым.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p110">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

13. <span data-ttu-id="e89cf-p111">Откройте окно браузера и перейдите по адресу `https://localhost:1337`, где находится интерфейс Vorlon.JS. Когда появится соответствующий запрос, выберите элемент **Всегда**, чтобы сделать сертификат безопасности доверенным.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p111">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface. When prompted, choose **Always** to trust the security certificate.</span></span>

    > [!NOTE]
    > <span data-ttu-id="e89cf-p112">Если запрос не появится, может потребоваться сделать сертификат доверенным вручную. Файл сертификата — `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Выполните указанные ниже действия. При возникновении проблем обратитесь к справке по Macintosh или iPad.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p112">If you are not prompted, you might need to trust the certificate manually. The certificate file is `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Try the following steps. If you have trouble, consult Macintosh or iPad help.</span></span>
    >
    > 1. <span data-ttu-id="e89cf-157">Закройте окно браузера, а затем в окне **Терминал**, где выполняется сервер Vorlon, нажмите клавиши CTRL+C, чтобы остановить сервер.</span><span class="sxs-lookup"><span data-stu-id="e89cf-157">Close the browser window and in the **Terminal** window that is running the Vorlon server, use Control-C to stop the server.</span></span>
    > 2. <span data-ttu-id="e89cf-p113">В средстве **Finder**, щелкните правой кнопкой мыши файл `server.crt` и выберите **Доступ к цепочкам ключей**. Откроется окно **Доступ к цепочкам ключей**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p113">In **Finder**, right-click the `server.crt` file and select **Keychain Access**. The **Keychain Access** window opens.</span></span>
    > 3. <span data-ttu-id="e89cf-p114">В списке **Цепочки ключей** слева выберите **Имя пользователя** (если оно еще не выбрано), а затем выберите **Сертификаты** в разделе **Категория**. В списке отображается сертификат **localhost**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p114">In the **Keychains** list on the left, select **login** if it is not already selected, and then select **Certificates** in the **Category** section. The certificate **localhost** is listed.</span></span>
    > 4. <span data-ttu-id="e89cf-p115">Щелкните сертификат **localhost** правой кнопкой мыши и выберите пункт **Получить сведения**. Откроется окно **localhost**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p115">Right-click the certificate **localhost** and select **Get Info**. A **localhost** window opens.</span></span>
    > 5. <span data-ttu-id="e89cf-164">В разделе **Доверие** откройте селектор **При использовании этого сертификата** и выберите **Всегда доверять**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-164">In the **Trust** section, open the selector labeled **When using this certificate** and select **Always Trust**.</span></span> 
    > 6. <span data-ttu-id="e89cf-p116">Закройте окно **localhost**. Если действие выполнено успешно, на значке сертификата **localhost** в окне **Доступ к цепочкам ключей** будет отображаться белый крест в синем круге.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p116">Close the **localhost** window. If the action was successful, the **localhost** certificate in the **Keychain Access** window has a white cross in a blue circle on its icon.</span></span>


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a><span data-ttu-id="e89cf-167">Конфигурация надстройки для отладки с помощью Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="e89cf-167">Configure the add-in for Vorlon.JS debugging</span></span>

1. <span data-ttu-id="e89cf-168">Добавьте следующий тег сценария в раздел `<head>` файла home.html (или основного HTML-файла) своей надстройки:</span><span class="sxs-lookup"><span data-stu-id="e89cf-168">Add the following script tag to the `<head>` section of the home.html file (or main HTML file) of your add-in:</span></span>

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>
    ```  

2. <span data-ttu-id="e89cf-169">Разверните веб-приложение надстройки на веб-сервере, доступном с Mac или iPad, например на веб-сайте Azure.</span><span class="sxs-lookup"><span data-stu-id="e89cf-169">Deploy the add-in web application to a web server that is accessible from the Mac or iPad, such as an Azure website.</span></span>

3. <span data-ttu-id="e89cf-170">Обновите URL-адрес надстройки во всех разделах манифеста, где он отображается.</span><span class="sxs-lookup"><span data-stu-id="e89cf-170">Update the URL of the add-in in all the places where the URL appears in the add-in manifest.</span></span>

4. <span data-ttu-id="e89cf-171">Скопируйте манифест надстройки в следующую папку на Mac или iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, где *{host_name}* — это Word, Excel, PowerPoint или Outlook.</span><span class="sxs-lookup"><span data-stu-id="e89cf-171">Copy the add-in manifest to the following folder on the Mac or iPad: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, where *{host_name}* is Word, Excel, PowerPoint, or Outlook.</span></span>


### <a name="inspect-an-add-in-in-vorlonjs"></a><span data-ttu-id="e89cf-172">Просмотр надстройки в Vorlon.JS</span><span class="sxs-lookup"><span data-stu-id="e89cf-172">Inspect an add-in in Vorlon.JS</span></span>

1. <span data-ttu-id="e89cf-173">Если сервер Vorlon не запущен, в средстве **Finder** перейдите к папке `/usr/local/lib/node_modules/vorlon`, щелкните правой кнопкой мыши вложенную папку `Server` и выберите пункт **Новый терминал в этой папке**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-173">If the Vorlon server is not running, in **Finder**, navigate to `/usr/local/lib/node_modules/vorlon`, right-click the `Server` subfolder, and select **New terminal at folder**.</span></span> 

2.  <span data-ttu-id="e89cf-p117">В окне **Терминал** введите команду `sudo vorlon`. Вам будет предложено ввести пароль администратора. Запустится сервер Vorlon. Оставьте окно **Терминал** открытым.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p117">In the **Terminal** window, enter `sudo vorlon`. You will be prompted to enter your administrator password. The Vorlon server starts. Leave the **Terminal** window open.</span></span>

3.  <span data-ttu-id="e89cf-178">Откройте окно браузера и перейдите по адресу `https://localhost:1337`, где находится интерфейс Vorlon.JS.</span><span class="sxs-lookup"><span data-stu-id="e89cf-178">Open a browser window and go to `https://localhost:1337`, which is the Vorlon.JS interface.</span></span>

4. <span data-ttu-id="e89cf-179">Загрузите неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="e89cf-179">Sideload the add-in.</span></span> <span data-ttu-id="e89cf-180">Если надстройка предназначена для Excel, PowerPoint или Word, загружайте ее согласно указаниям из статьи [Загрузка неопубликованных надстроек Office на iPad и компьютер Mac](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="e89cf-180">If it is for Excel, PowerPoint, or Word, sideload it as described in [Sideload an Office Add-in on iPad and Mac](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="e89cf-181">Если же это надстройка Outlook, загружайте ее согласно указаниям из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="e89cf-181">If it is an Outlook add-in, sideload it as described in [Sideload Outlook Add-ins for testing](https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span></span> <span data-ttu-id="e89cf-182">Если надстройка не использует команды, она откроется сразу.</span><span class="sxs-lookup"><span data-stu-id="e89cf-182">If the add-in does not use add-in commands, it will open immediately.</span></span> <span data-ttu-id="e89cf-183">В противном случае нажмите кнопку, чтобы открыть надстройку.</span><span class="sxs-lookup"><span data-stu-id="e89cf-183">Otherwise, choose the button to open the add-in.</span></span> <span data-ttu-id="e89cf-184">В зависимости от сборки ведущего приложения Office, кнопка будет отображаться на вкладке **Главная** или **Надстройка**.</span><span class="sxs-lookup"><span data-stu-id="e89cf-184">Depending on the build of the Office host application, the button will be on either the **Home** tab or an **Add-in** tab.</span></span>

<span data-ttu-id="e89cf-185">Надстройка будет отображаться в списке клиентов в Vorlon.JS (в левой части интерфейса Vorlon.JS) как **{ОС} - n**, где *n* — число, а *{ОС}* — тип устройства, например Macintosh.</span><span class="sxs-lookup"><span data-stu-id="e89cf-185">The add-in will show up in the list of Clients in Vorlon.JS (on the left side of the Vorlon.JS interface) as **{OS} - n**, for some number *n*, and where *{OS}* is the device type, such as "Macintosh".</span></span>

![Снимок экрана с интерфейсом Vorlon.js](../images/vorlon-interface.png)

<span data-ttu-id="e89cf-187">Для средства Vorlon доступен ряд подключаемых модулей. Те, что уже включены, отображаются в виде вкладок в верхней части окна.</span><span class="sxs-lookup"><span data-stu-id="e89cf-187">The Vorlon tool has a variety of plug-ins. The ones that are currently enabled appear as tabs at the top of the tool.</span></span> <span data-ttu-id="e89cf-188">(Вы можете включить дополнительные подключаемые модули, выбрав значок шестеренки слева.) Эти подключаемые модули подобны функциям, вызываемым с помощью клавиши F12.</span><span class="sxs-lookup"><span data-stu-id="e89cf-188">(You can enable more plug-ins by choosing the gears icon on the left.) These plug-ins are  similar to the functions in F12 tools.</span></span> <span data-ttu-id="e89cf-189">Например, вы можете выделять элементы модели DOM, выполнять команды и другие действия.</span><span class="sxs-lookup"><span data-stu-id="e89cf-189">For example, you can highlight DOM elements, execute commands, and more.</span></span> <span data-ttu-id="e89cf-190">Дополнительные сведения см. в разделе с описанием [основных подключаемых модулей](http://vorlonjs.com/documentation/#console) документации по Vorlon.</span><span class="sxs-lookup"><span data-stu-id="e89cf-190">For more details, see [Vorlon Documentation Core Plugins](http://vorlonjs.com/documentation/#console)</span></span>

<span data-ttu-id="e89cf-p120">Подключаемый модуль **Надстройка Office** добавляет в Office.js такие возможности, как изучение объектной модели, совершение вызовов Office.js и считывание значений свойств объекта. Сведения см. в статье [Подключаемый модуль VorlonJS для отладки надстроек Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span><span class="sxs-lookup"><span data-stu-id="e89cf-p120">An **Office Addin** plug-in adds extra capabilities for Office.js, such as exploring the object model, executing Office.js calls, and reading the values of object properties. For instructions, see [VorlonJS plugin for debugging Office Add-in](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).</span></span>

> [!NOTE]
> <span data-ttu-id="e89cf-193">В Vorlon.JS невозможно задавать точки останова.</span><span class="sxs-lookup"><span data-stu-id="e89cf-193">There is no way to set break points in Vorlon.JS.</span></span>


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a><span data-ttu-id="e89cf-194">Очистка кэша приложения Office на компьютере Mac или iPad</span><span class="sxs-lookup"><span data-stu-id="e89cf-194">Clearing the Office application's cache on a Mac or iPad</span></span>

<span data-ttu-id="e89cf-p121">Для повышения производительности надстройки часто кэшируются в Office для Mac. Как правило, для очистки кэша необходимо перезагрузить надстройку. Если в одном документе несколько надстроек, автоматическая очистка кэша может не сработать при перезагрузке.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p121">Add-ins are cached often in Office for Mac, for performance reasons. Normally, the cache is cleared by reloading the add-in. If  more than one add-in exists in the same document, the process of automatically clearing the cache on reload might not be reliable.</span></span>

<span data-ttu-id="e89cf-198">На компьютере Mac можно очистить кэш вручную, удалив все содержимое папки `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="e89cf-198">On a Mac, you can clear the cache manually by deleting everything in the `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/` folder.</span></span>

<span data-ttu-id="e89cf-p122">На iPad можно вызвать в надстройке метод JavaScript `window.location.reload(true)` для принудительной перезагрузки. Вы также можете переустановить Office.</span><span class="sxs-lookup"><span data-stu-id="e89cf-p122">On an iPad, you can call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>
