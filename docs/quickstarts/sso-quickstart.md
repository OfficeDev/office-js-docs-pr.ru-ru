---
title: Создание надстройки Office, в которой используется единый вход (предварительная версия), с помощью генератора Yeoman
description: Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия) с помощью генератора Yeoman.
ms.date: 01/27/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: d3a78a99574c92d0066003f0e39e835563f473cd
ms.sourcegitcommit: 413f163729183994de61a8281685184b377ef76c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/28/2020
ms.locfileid: "41571403"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="1a3d0-103">Создание надстройки Office, в которой используется единый вход (предварительная версия), с помощью генератора Yeoman</span><span class="sxs-lookup"><span data-stu-id="1a3d0-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="1a3d0-104">В этой статье описывается процесс использования генератора Yeoman для создания надстройки Office для Excel, Outlook, Word или PowerPoint, в которой используется единый вход (SSO), когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Outlook, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="1a3d0-105">Прежде чем приступить к работе, познакомьтесь с основными понятиями, связанными с использованием единого входа в надстройках Office, с помощью статьи [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="1a3d0-106">Генератор Yeoman упрощает процесс создания надстройки с использованием единого входа, автоматизируя действия, необходимые для настройки единого входа в Azure, и создавая код, необходимый для его использования в надстройке.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="1a3d0-107">Подробное пошаговое руководство, в котором объясняется, как вручную выполнить действия, автоматизируемые генератором Yeoman, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1a3d0-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="1a3d0-108">Prerequisites</span></span>

* <span data-ttu-id="1a3d0-109">[Node.js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="1a3d0-109">[Node.js](https://nodejs.org) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

* <span data-ttu-id="1a3d0-110">Последняя версия [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="1a3d0-111">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="1a3d0-111">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="1a3d0-112">С помощью генератора Yeoman можно создать надстройку Office с использованием единого входа для Excel, Outlook, Word или PowerPoint, работа которой основана на сценарии JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-112">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Outlook, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="1a3d0-113">В приведенных ниже инструкциях указаны `JavaScript` и `Excel`, однако следует выбрать тип сценария и клиентское приложение Office, которое лучше всего подходит для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-113">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="1a3d0-114">**Выберите тип проекта:** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="1a3d0-114">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="1a3d0-115">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="1a3d0-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="1a3d0-116">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="1a3d0-116">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="1a3d0-117">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="1a3d0-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="1a3d0-119">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="1a3d0-120">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="1a3d0-120">Explore the project</span></span>

<span data-ttu-id="1a3d0-121">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит код для надстройки области задач с использованием единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-121">The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.</span></span>

- <span data-ttu-id="1a3d0-122">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-122">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="1a3d0-123">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-123">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="1a3d0-124">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-124">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="1a3d0-125">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-125">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="1a3d0-126">Файл **./src/helpers/documentHelper.js**использует библиотеку Office JavaScript для добавления данных из Microsoft Graph в документ Office.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-126">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="1a3d0-127">Файл **./src/helpers/fallbackauthdialog.html** — это страница без пользовательского интерфейса, которая загружает JavaScript резервного метода проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-127">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="1a3d0-128">Файл **./src/helpers/fallbackauthdialog.js** содержит сценарий JavaScript резервного метода проверки подлинности, который выполняется во время входа пользователя с помощью MSAL.js.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-128">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="1a3d0-129">Файл **./src/helpers/fallbackauthhelper.js** содержит JavaScript области задач, вызывающий резервный метод проверки подлинности при выполнении сценариев, если проверка подлинности на основе единого входа не поддерживается. </span><span class="sxs-lookup"><span data-stu-id="1a3d0-129">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="1a3d0-130">Файл **./src/helpers/ssoauthhelper.js** содержит вызов JavaScript для API единого входа, `getAccessToken`, получает маркер начальной загрузки, инициирует его замену на маркер доступа для Microsoft Graph и вызывает данные Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-130">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="1a3d0-131">Файл **./ENV** в корневом каталоге проекта определяет константы, используемые в проекте надстройки. </span><span class="sxs-lookup"><span data-stu-id="1a3d0-131">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="1a3d0-132">Некоторые константы, определяемые в этом файле, используются для упрощения процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-132">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="1a3d0-133">Вам может потребоваться обновить значения в этом файле в соответствии с конкретным сценарием.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-133">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="1a3d0-134">Например, вы можете обновить значение области, если для надстройки требуется не `User.Read`, а другое разрешение.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-134">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="1a3d0-135">Настройка единого входа</span><span class="sxs-lookup"><span data-stu-id="1a3d0-135">Configure SSO</span></span>

<span data-ttu-id="1a3d0-136">На этом этапе проект надстройки уже создан и содержит код, необходимый для упрощения процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-136">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="1a3d0-137">Выполните указанные ниже действия, чтобы настроить единый вход для вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-137">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="1a3d0-138">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-138">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="1a3d0-139">Чтобы настроить единый вход для надстройки, выполните приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-139">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="1a3d0-140">Эта команда приведет к ошибке, если для клиента настроена двухфакторная проверка подлинности.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-140">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="1a3d0-141">В этом случае вам потребуется выполнить регистрацию приложения в Azure и настройку единого входа вручную, как описано в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-141">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="1a3d0-142">Откроется окно веб-браузера, в котором вам будет предложено войти в Azure.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-142">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="1a3d0-143">Войдите в Azure, используя учетные данные администратора Office 365.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-143">Sign in to Azure using your Office 365 administrator credentials.</span></span> <span data-ttu-id="1a3d0-144">Эти учетные данные будут использоваться для регистрации нового приложения в Azure и настройки параметров, необходимых для единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-144">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1a3d0-145">Если на этом этапе для входа в Azure вы используете учетные данные без прав администратора, сценарий `configure-sso` не сможет предоставить согласие администратора для надстройки пользователям в организации.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-145">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="1a3d0-146">В этом случае единый вход будет недоступен для пользователей надстройки, и им будет предложено выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-146">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="1a3d0-147">После ввода учетных данных закройте окно браузера и вернитесь к командной строке.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-147">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="1a3d0-148">В процессе настройки единого входа на консоль будут выводиться сообщения о состоянии.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-148">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="1a3d0-149">В соответствии с ними, файлы проекта надстройки, созданные генератором Yeoman, автоматически обновляются с учетом данных, необходимых для процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-149">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="1a3d0-150">Проверка</span><span class="sxs-lookup"><span data-stu-id="1a3d0-150">Try it out</span></span>

<span data-ttu-id="1a3d0-151">Если вы создали надстройку Excel, Word или PowerPoint, выполните действия из следующего раздела, чтобы попробовать ее в действии. Если вы создали надстройку Outlook, выполните действия из раздела [Outlook](#outlook).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-151">If you've created an Excel, Word, or PowerPoint add-in, complete the steps in the following section to try it out. If you've created an Outlook add-in, complete the steps in the [Outlook](#outlook) section instead.</span></span>

### <a name="excel-word-and-powerpoint"></a><span data-ttu-id="1a3d0-152">Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="1a3d0-152">Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="1a3d0-153">Выполните следующие действия, чтобы испытать надстройку Excel, Word или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-153">Complete the following steps to try out an Excel, Word, or PowerPoint add-in.</span></span>

1. <span data-ttu-id="1a3d0-154">Когда процесс настройки единого входа будет завершен, для построения проекта, запуска локального веб-сервера и загрузки своей надстройки в ранее выбранное клиентское приложение Office запустите указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-154">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1a3d0-155">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-155">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="1a3d0-156">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-156">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="1a3d0-157">Убедитесь, что в клиентском приложении Office (например, Excel, Word или PowerPoint), которое открывается при запуске указанной выше команды, вы выполнили вход как участник той же организации Office 365, что и администратор, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на этапе 3, описанном в [предыдущем разделе](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-157">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="1a3d0-158">Благодаря этому будут созданы соответствующие условия для успешного единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-158">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="1a3d0-159">В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-159">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="1a3d0-160">На рисунке ниже показана эта кнопка в Excel. </span><span class="sxs-lookup"><span data-stu-id="1a3d0-160">The following image shows this button in Excel.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="1a3d0-162">В нижней части области задач нажмите кнопку **Получить сведения о моем профиле пользователя**, чтобы начать процесс единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-162">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

5. <span data-ttu-id="1a3d0-163">Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-163">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="1a3d0-164">Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-164">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="1a3d0-165">Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-165">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="1a3d0-167">После принятия пользователем запрос разрешений больше не выводится на экран.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-167">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="1a3d0-168">Надстройка получает сведения о профиле пользователя, выполнившего вход, и вносит их в документ.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-168">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="1a3d0-169">На приведенном ниже рисунке показан пример сведений о профиле, внесенных на лист Excel.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-169">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Сведения о профиле пользователя на листе Excel](../images/sso-user-profile-info-excel.png)

### <a name="outlook"></a><span data-ttu-id="1a3d0-171">Outlook</span><span class="sxs-lookup"><span data-stu-id="1a3d0-171">Outlook</span></span>

<span data-ttu-id="1a3d0-172">Выполните следующие действия, чтобы испытать надстройку Outlook.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-172">Complete the following steps to try out an Outlook add-in.</span></span>

1. <span data-ttu-id="1a3d0-173">По завершении процесса настройки единого входа выполните следующую команду, чтобы создать проект и запустить локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-173">When the SSO configuration process completes, run the following command to build the project and start the local web server.</span></span>

    > [!NOTE]
    > <span data-ttu-id="1a3d0-174">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-174">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="1a3d0-175">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-175">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="1a3d0-176">Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](/outlook/add-ins/sideload-outlook-add-ins-for-testing).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-176">Follow the instructions in [Sideload Outlook add-ins for testing](/outlook/add-ins/sideload-outlook-add-ins-for-testing) to sideload the add-in in Outlook.</span></span> <span data-ttu-id="1a3d0-177">Убедитесь, что вход в Outlook выполнен в качестве участника той же организации Office 365, что и администратор, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на этапе 3, описанном в [предыдущем разделе](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-177">Make sure that you're signed in to Outlook with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="1a3d0-178">Благодаря этому будут созданы соответствующие условия для успешного единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-178">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="1a3d0-179">В Outlook создайте новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-179">In Outlook, compose a new message.</span></span>

4. <span data-ttu-id="1a3d0-180">В окне создания сообщения нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-180">In the message compose window, choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Outlook](../images/outlook-sso-ribbon-button.png)

5. <span data-ttu-id="1a3d0-182">В нижней части области задач нажмите кнопку **Получить сведения о моем профиле пользователя**, чтобы начать процесс единого входа.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-182">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

6. <span data-ttu-id="1a3d0-183">Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-183">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="1a3d0-184">Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-184">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="1a3d0-185">Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-185">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="1a3d0-187">После принятия пользователем запрос разрешений больше не выводится на экран.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-187">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

7. <span data-ttu-id="1a3d0-188">Надстройка получает сведения о профиле пользователя, выполнившего вход, и вносит их в текст сообщения электронной почты.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-188">The add-in retrieves profile information for the signed-in user and writes it to the body of the email message.</span></span> 

    ![Сведения о профиле пользователя в сообщении Outlook](../images/sso-user-profile-info-outlook.png)

## <a name="next-steps"></a><span data-ttu-id="1a3d0-190">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="1a3d0-190">Next steps</span></span>

<span data-ttu-id="1a3d0-191">Поздравляем! Вы успешно создали надстройку области задач, в которой используется единый вход, когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="1a3d0-191">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="1a3d0-192">Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="1a3d0-192">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="1a3d0-193">См. также</span><span class="sxs-lookup"><span data-stu-id="1a3d0-193">See also</span></span>

- [<span data-ttu-id="1a3d0-194">Включение единого входа для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1a3d0-194">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="1a3d0-195">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="1a3d0-195">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="1a3d0-196">Устранение ошибок единого входа</span><span class="sxs-lookup"><span data-stu-id="1a3d0-196">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)