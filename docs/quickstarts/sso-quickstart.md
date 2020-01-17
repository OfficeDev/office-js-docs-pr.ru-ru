---
title: Создание надстройки Office, в которой используется единый вход (предварительная версия), с помощью генератора Yeoman
description: Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия) с помощью генератора Yeoman.
ms.date: 01/13/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 1f02f03fec0d6be32fc7a0d6b98fce30e19c28e2
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217367"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="fef6b-103">Создание надстройки Office, в которой используется единый вход (предварительная версия), с помощью генератора Yeoman</span><span class="sxs-lookup"><span data-stu-id="fef6b-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="fef6b-104">В этой статье описывается процесс использования генератора Yeoman для создания надстройки Office для Excel, Word или PowerPoint, в которой используется единый вход (SSO), когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fef6b-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="fef6b-105">Прежде чем приступить к работе, познакомьтесь с основными понятиями, связанными с использованием единого входа в надстройках Office, с помощью статьи [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="fef6b-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="fef6b-106">Генератор Yeoman упрощает процесс создания надстройки с использованием единого входа, автоматизируя действия, необходимые для настройки единого входа в Azure, и создавая код, необходимый для его использования в надстройке.</span><span class="sxs-lookup"><span data-stu-id="fef6b-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="fef6b-107">Подробное пошаговое руководство, в котором объясняется, как вручную выполнить действия, автоматизируемые генератором Yeoman, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="fef6b-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="fef6b-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="fef6b-108">Prerequisites</span></span>

* <span data-ttu-id="fef6b-109">[Node.js](https://nodejs.org) (версия 10.15.0 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="fef6b-109">[Node.js](https://nodejs.org) (version 10.15.0 or later)</span></span>

* <span data-ttu-id="fef6b-110">Последняя версия [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="fef6b-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="fef6b-111">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="fef6b-111">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="fef6b-112">С помощью генератора Yeoman можно создать надстройку Office с использованием единого входа для Excel, Word или PowerPoint, работа которой основана на сценарии JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="fef6b-112">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="fef6b-113">В приведенных ниже инструкциях указаны `JavaScript` и `Excel`, однако следует выбрать тип сценария и клиентское приложение Office, которое лучше всего подходит для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="fef6b-113">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="fef6b-114">**Выберите тип проекта:** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="fef6b-114">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="fef6b-115">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="fef6b-115">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="fef6b-116">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="fef6b-116">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="fef6b-117">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="fef6b-117">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="fef6b-119">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="fef6b-119">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="fef6b-120">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="fef6b-120">Explore the project</span></span>

<span data-ttu-id="fef6b-121">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит код для надстройки области задач с использованием единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-121">The add-in project that you've created with the Yeoman generator contains code for an SSO-enabled task pane add-in.</span></span>

- <span data-ttu-id="fef6b-122">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="fef6b-122">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="fef6b-123">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="fef6b-123">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="fef6b-124">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="fef6b-124">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="fef6b-125">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="fef6b-125">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="fef6b-126">Файл **./src/helpers/documentHelper.js**использует библиотеку Office JavaScript для добавления данных из Microsoft Graph в документ Office.</span><span class="sxs-lookup"><span data-stu-id="fef6b-126">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="fef6b-127">Файл **./src/helpers/fallbackauthdialog.html** — это страница без пользовательского интерфейса, которая загружает JavaScript резервного метода проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="fef6b-127">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="fef6b-128">Файл **./src/helpers/fallbackauthdialog.js** содержит сценарий JavaScript резервного метода проверки подлинности, который выполняется во время входа пользователя с помощью MSAL.js.</span><span class="sxs-lookup"><span data-stu-id="fef6b-128">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="fef6b-129">Файл **./src/helpers/fallbackauthhelper.js** содержит JavaScript области задач, вызывающий резервный метод проверки подлинности при выполнении сценариев, если проверка подлинности на основе единого входа не поддерживается. </span><span class="sxs-lookup"><span data-stu-id="fef6b-129">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="fef6b-130">Файл **./src/helpers/ssoauthhelper.js** содержит вызов JavaScript для API единого входа, `getAccessToken`, получает маркер начальной загрузки, инициирует его замену на маркер доступа для Microsoft Graph и вызывает данные Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="fef6b-130">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="fef6b-131">Файл **./ENV** в корневом каталоге проекта определяет константы, используемые в проекте надстройки. </span><span class="sxs-lookup"><span data-stu-id="fef6b-131">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="fef6b-132">Некоторые константы, определяемые в этом файле, используются для упрощения процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-132">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="fef6b-133">Вам может потребоваться обновить значения в этом файле в соответствии с конкретным сценарием.</span><span class="sxs-lookup"><span data-stu-id="fef6b-133">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="fef6b-134">Например, вы можете обновить значение области, если для надстройки требуется не `User.Read`, а другое разрешение.</span><span class="sxs-lookup"><span data-stu-id="fef6b-134">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="fef6b-135">Настройка единого входа</span><span class="sxs-lookup"><span data-stu-id="fef6b-135">Configure SSO</span></span>

<span data-ttu-id="fef6b-136">На этом этапе проект надстройки уже создан и содержит код, необходимый для упрощения процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-136">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="fef6b-137">Выполните указанные ниже действия, чтобы настроить единый вход для вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="fef6b-137">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="fef6b-138">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="fef6b-138">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="fef6b-139">Чтобы настроить единый вход для надстройки, выполните приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="fef6b-139">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="fef6b-140">Эта команда приведет к ошибке, если для клиента настроена двухфакторная проверка подлинности.</span><span class="sxs-lookup"><span data-stu-id="fef6b-140">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="fef6b-141">В этом случае вам потребуется выполнить регистрацию приложения в Azure и настройку единого входа вручную, как описано в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="fef6b-141">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="fef6b-142">Откроется окно веб-браузера, в котором вам будет предложено войти в Azure.</span><span class="sxs-lookup"><span data-stu-id="fef6b-142">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="fef6b-143">Войдите в Azure, используя учетные данные администратора Office 365.</span><span class="sxs-lookup"><span data-stu-id="fef6b-143">Sign in to Azure using your Office 365 administrator credentials.</span></span> <span data-ttu-id="fef6b-144">Эти учетные данные будут использоваться для регистрации нового приложения в Azure и настройки параметров, необходимых для единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-144">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="fef6b-145">Если на этом этапе для входа в Azure вы используете учетные данные без прав администратора, сценарий `configure-sso` не сможет предоставить согласие администратора для надстройки пользователям в организации.</span><span class="sxs-lookup"><span data-stu-id="fef6b-145">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="fef6b-146">В этом случае единый вход будет недоступен для пользователей надстройки, и им будет предложено выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="fef6b-146">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="fef6b-147">После ввода учетных данных закройте окно браузера и вернитесь к командной строке.</span><span class="sxs-lookup"><span data-stu-id="fef6b-147">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="fef6b-148">В процессе настройки единого входа на консоль будут выводиться сообщения о состоянии.</span><span class="sxs-lookup"><span data-stu-id="fef6b-148">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="fef6b-149">В соответствии с ними, файлы проекта надстройки, созданные генератором Yeoman, автоматически обновляются с учетом данных, необходимых для процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-149">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="fef6b-150">Попробуйте</span><span class="sxs-lookup"><span data-stu-id="fef6b-150">Try it out</span></span>

1. <span data-ttu-id="fef6b-151">Когда процесс настройки единого входа будет завершен, для построения проекта, запуска локального веб-сервера и загрузки своей надстройки в ранее выбранное клиентское приложение Office запустите указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="fef6b-151">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="fef6b-152">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="fef6b-152">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="fef6b-153">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="fef6b-153">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="fef6b-154">Убедитесь, что в клиентском приложении Office (например, Excel, Word или PowerPoint), которое открывается при запуске указанной выше команды, вы выполнили вход как участник той же организации Office 365, что и администратор, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на этапе 3, описанном в [предыдущем разделе](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="fef6b-154">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="fef6b-155">Благодаря этому будут созданы соответствующие условия для успешного единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-155">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="fef6b-156">В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="fef6b-156">In the Office client application, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="fef6b-157">На рисунке ниже показана эта кнопка в Excel. </span><span class="sxs-lookup"><span data-stu-id="fef6b-157">The following image shows this button in Excel.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="fef6b-159">В нижней части области задач нажмите кнопку **Получить сведения о моем профиле пользователя**, чтобы начать процесс единого входа.</span><span class="sxs-lookup"><span data-stu-id="fef6b-159">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

    > [!NOTE] 
    > <span data-ttu-id="fef6b-160">Если на этом этапе вы еще не вошли в Office, вам будет предложено сделать это.</span><span class="sxs-lookup"><span data-stu-id="fef6b-160">If you're not already signed in to Office at this point, you'll be prompted to sign in.</span></span> <span data-ttu-id="fef6b-161">Как говорилось выше, вам нужно выполнить вход в качестве участника той же организации Office 365, что и администратор, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на этапе 3, описанном в [предыдущем разделе](#configure-sso). Это необходимо для успешного единого входа. </span><span class="sxs-lookup"><span data-stu-id="fef6b-161">As described previously, you should sign in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso), if you want SSO to succeed.</span></span>

5. <span data-ttu-id="fef6b-162">Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя.</span><span class="sxs-lookup"><span data-stu-id="fef6b-162">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="fef6b-163">Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи).</span><span class="sxs-lookup"><span data-stu-id="fef6b-163">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="fef6b-164">Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="fef6b-164">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="fef6b-166">После принятия пользователем запрос разрешений больше не выводится на экран.</span><span class="sxs-lookup"><span data-stu-id="fef6b-166">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="fef6b-167">Надстройка получает сведения о профиле пользователя, выполнившего вход, и вносит их в документ.</span><span class="sxs-lookup"><span data-stu-id="fef6b-167">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="fef6b-168">На приведенном ниже рисунке показан пример сведений о профиле, внесенных на лист Excel.</span><span class="sxs-lookup"><span data-stu-id="fef6b-168">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Сведения о профиле пользователя на листе Excel](../images/sso-user-profile-info-excel.png)

## <a name="next-steps"></a><span data-ttu-id="fef6b-170">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="fef6b-170">Next steps</span></span>

<span data-ttu-id="fef6b-171">Поздравляем! Вы успешно создали надстройку области задач, в которой используется единый вход, когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="fef6b-171">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="fef6b-172">Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="fef6b-172">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="fef6b-173">См. также</span><span class="sxs-lookup"><span data-stu-id="fef6b-173">See also</span></span>

- [<span data-ttu-id="fef6b-174">Включение единого входа для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="fef6b-174">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="fef6b-175">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="fef6b-175">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="fef6b-176">Устранение ошибок единого входа</span><span class="sxs-lookup"><span data-stu-id="fef6b-176">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)