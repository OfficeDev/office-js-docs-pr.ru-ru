---
title: Создание надстройки Office, в которой используется единый вход (предварительная версия), с помощью генератора Yeoman
description: Создание надстройки Office на платформе Node.js с использованием единого входа (предварительная версия) с помощью генератора Yeoman.
ms.date: 01/13/2020
ms.prod: non-product-specific
localization_priority: Priority
ms.openlocfilehash: 3c67fdb2b8582546c13624dcb8a6f139bb638df0
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111120"
---
# <a name="use-the-yeoman-generator-to-create-an-office-add-in-that-uses-single-sign-on-preview"></a><span data-ttu-id="a5995-103">Создание надстройки Office, в которой используется единый вход (предварительная версия), с помощью генератора Yeoman</span><span class="sxs-lookup"><span data-stu-id="a5995-103">Use the Yeoman generator to create an Office Add-in that uses single sign-on (preview)</span></span>

<span data-ttu-id="a5995-104">В этой статье описывается процесс использования генератора Yeoman для создания надстройки Office для Excel, Word или PowerPoint, в которой используется единый вход (SSO), когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a5995-104">In this article, you'll walk through the process of using the Yeoman generator to create an Office Add-in for Excel, Word, or PowerPoint that uses single sign-on (SSO) when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span>

> [!TIP]
> <span data-ttu-id="a5995-105">Прежде чем приступить к работе, познакомьтесь с основными понятиями, связанными с использованием единого входа в надстройках Office, с помощью статьи [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="a5995-105">Before you attempt to complete this quick start, review [Enable single sign-on for Office Add-ins](../develop/sso-in-office-add-ins.md) to learn basic concepts about SSO in Office Add-ins.</span></span> 
 
<span data-ttu-id="a5995-106">Генератор Yeoman упрощает процесс создания надстройки с использованием единого входа, автоматизируя действия, необходимые для настройки единого входа в Azure, и создавая код, необходимый для его использования в надстройке.</span><span class="sxs-lookup"><span data-stu-id="a5995-106">The Yeoman generator simplifies the process of creating an SSO add-in, by automating the steps required to configure SSO within Azure and generating the code that's necessary for an add-in to use SSO.</span></span> <span data-ttu-id="a5995-107">Подробное пошаговое руководство, в котором объясняется, как вручную выполнить действия, автоматизируемые генератором Yeoman, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="a5995-107">For a detailed walkthrough that describes how to manually complete the steps that the Yeoman generator automates, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a5995-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="a5995-108">Prerequisites</span></span>

- <span data-ttu-id="a5995-109">[Node.js](https://nodejs.org) (версия 10.15.0 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="a5995-109">[Node.js](https://nodejs.org) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="a5995-110">Последняя версия [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="a5995-110">The latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```command&nbsp;line
    npm install -g yo generator-office
    ```

    [!include[note to update Yeoman generator](../includes/note-yeoman-generator-update.md)]

- <span data-ttu-id="a5995-111">Учетная запись Office 365 (версии Office, предоставляемые по подписке).</span><span class="sxs-lookup"><span data-stu-id="a5995-111">Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="a5995-112">Если у вас еще нет учетной записи Office 365, вы можете оформить бесплатную возобновляемую подписку на Office 365 на 90 дней, присоединившись к [программе для разработчиков Office 365](https://aka.ms/devprogramsignup).</span><span class="sxs-lookup"><span data-stu-id="a5995-112">If you don't already have an Office 365 account, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://aka.ms/devprogramsignup).</span></span> 

- <span data-ttu-id="a5995-113">Сборка Office 365 для участников программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="a5995-113">An Insider's build of Office 365.</span></span> <span data-ttu-id="a5995-114">Чтобы получить эту версию, следует использовать последнюю ежемесячную версию и сборку из канала программы предварительной оценки, при этом необходимо [быть участником программы предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1).</span><span class="sxs-lookup"><span data-stu-id="a5995-114">You should use the latest monthly version and build from the Insiders channel but you need to be an Office Insider to get this version.</span></span> 

    > [!NOTE]
    > <span data-ttu-id="a5995-115">Когда сборка будет готова к выпуску на канале Semi-Annual Channel, для нее будет отключена поддержка функций предварительной версии, включая единый вход.</span><span class="sxs-lookup"><span data-stu-id="a5995-115">Please note that when a build graduates to the production semi-annual channel, support for preview features, including SSO, is turned off for that build.</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="a5995-116">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="a5995-116">Create the add-in project</span></span>

> [!TIP]
> <span data-ttu-id="a5995-117">С помощью генератора Yeoman можно создать надстройку Office с использованием единого входа для Excel, Word или PowerPoint, работа которой основана на сценарии JavaScript или TypeScript.</span><span class="sxs-lookup"><span data-stu-id="a5995-117">The Yeoman generator can create an SSO-enabled Office Add-in for Excel, Word, or PowerPoint, and can be created with script type of JavaScript or TypeScript.</span></span> <span data-ttu-id="a5995-118">В приведенных ниже инструкциях указаны `JavaScript` и `Excel`, однако следует выбрать тип сценария и клиентское приложение Office, которое лучше всего подходит для вашего сценария.</span><span class="sxs-lookup"><span data-stu-id="a5995-118">The following instructions specify `JavaScript` and `Excel`, but you should choose the script type and Office client application that best suits your scenario.</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="a5995-119">**Выберите тип проекта:** `Office Add-in Task Pane project supporting single sign-on`</span><span class="sxs-lookup"><span data-stu-id="a5995-119">**Choose a project type:** `Office Add-in Task Pane project supporting single sign-on`</span></span>
- <span data-ttu-id="a5995-120">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="a5995-120">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="a5995-121">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="a5995-121">**What do you want to name your add-in?**</span></span> `My SSO Office Add-in`
- <span data-ttu-id="a5995-122">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="a5995-122">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-sso-excel.png)

<span data-ttu-id="a5995-124">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="a5995-124">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="a5995-125">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="a5995-125">Explore the project</span></span>

<span data-ttu-id="a5995-126">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит код для надстройки области задач с использованием единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-126">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span>

- <span data-ttu-id="a5995-127">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5995-127">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="a5995-128">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="a5995-128">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="a5995-129">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="a5995-129">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="a5995-130">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="a5995-130">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

- <span data-ttu-id="a5995-131">Файл **./src/helpers/documentHelper.js**использует библиотеку Office JavaScript для добавления данных из Microsoft Graph в документ Office.</span><span class="sxs-lookup"><span data-stu-id="a5995-131">The **./src/helpers/documentHelper.js** file uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>
- <span data-ttu-id="a5995-132">Файл **./src/helpers/fallbackauthdialog.html** — это страница без пользовательского интерфейса, которая загружает JavaScript резервного метода проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="a5995-132">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the fallback authentication method's JavaScript.</span></span>
- <span data-ttu-id="a5995-133">Файл **./src/helpers/fallbackauthdialog.js** содержит сценарий JavaScript резервного метода проверки подлинности, который выполняется во время входа пользователя с помощью MSAL.js.</span><span class="sxs-lookup"><span data-stu-id="a5995-133">The **./src/helpers/fallbackauthdialog.js** file contains the fallback authentication method's JavaScript that signs on the user with msal.js.</span></span>
- <span data-ttu-id="a5995-134">Файл **./src/helpers/fallbackauthhelper.js** содержит JavaScript области задач, вызывающий резервный метод проверки подлинности при выполнении сценариев, если проверка подлинности на основе единого входа не поддерживается. </span><span class="sxs-lookup"><span data-stu-id="a5995-134">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication method in scenarios when SSO authentication is not supported.</span></span>
- <span data-ttu-id="a5995-135">Файл **./src/helpers/ssoauthhelper.js** содержит вызов JavaScript для API единого входа, `getAccessToken`, получает маркер начальной загрузки, инициирует его замену на маркер доступа для Microsoft Graph и вызывает данные Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="a5995-135">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>

- <span data-ttu-id="a5995-136">Файл **./ENV** в корневом каталоге проекта определяет константы, используемые в проекте надстройки. </span><span class="sxs-lookup"><span data-stu-id="a5995-136">The **./ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>
    > [!NOTE]
    > <span data-ttu-id="a5995-137">Некоторые константы, определяемые в этом файле, используются для упрощения процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-137">Some of the constants defined in this file are used to facilitate the SSO process.</span></span> <span data-ttu-id="a5995-138">Вам может потребоваться обновить значения в этом файле в соответствии с конкретным сценарием.</span><span class="sxs-lookup"><span data-stu-id="a5995-138">You may want to update values in this file to match your specific scenario.</span></span> <span data-ttu-id="a5995-139">Например, вы можете обновить значение области, если для надстройки требуется не `User.Read`, а другое разрешение.</span><span class="sxs-lookup"><span data-stu-id="a5995-139">For example, you can update this file to specify a different scope, if your add-in requires something other than `User.Read`.</span></span>

## <a name="configure-sso"></a><span data-ttu-id="a5995-140">Настройка единого входа</span><span class="sxs-lookup"><span data-stu-id="a5995-140">Configure SSO</span></span>

<span data-ttu-id="a5995-141">На этом этапе проект надстройки уже создан и содержит код, необходимый для упрощения процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-141">At this point, your add-in project has been created and contains the code that's necessary to facilitate the SSO process.</span></span> <span data-ttu-id="a5995-142">Выполните указанные ниже действия, чтобы настроить единый вход для вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5995-142">Next, complete the following steps to configure SSO for your add-in.</span></span>

1. <span data-ttu-id="a5995-143">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="a5995-143">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My SSO Office Add-in"
    ```

2. <span data-ttu-id="a5995-144">Чтобы настроить единый вход для надстройки, выполните приведенную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="a5995-144">Run the following command to configure SSO for the add-in.</span></span>

    ```command&nbsp;line
    npm run configure-sso
    ```

    > [!WARNING]
    > <span data-ttu-id="a5995-145">Эта команда приведет к ошибке, если для клиента настроена двухфакторная проверка подлинности.</span><span class="sxs-lookup"><span data-stu-id="a5995-145">This command will fail if your tenant is configured to require two-factor authentication.</span></span> <span data-ttu-id="a5995-146">В этом случае вам потребуется выполнить регистрацию приложения в Azure и настройку единого входа вручную, как описано в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="a5995-146">In this scenario, you'll need to manually complete the Azure app registration and SSO configuration steps, as described in the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

3. <span data-ttu-id="a5995-147">Откроется окно веб-браузера, в котором вам будет предложено войти в Azure.</span><span class="sxs-lookup"><span data-stu-id="a5995-147">A web browser window will open and prompt you to sign in to Azure.</span></span> <span data-ttu-id="a5995-148">Войдите в Azure, используя учетные данные администратора Office 365.</span><span class="sxs-lookup"><span data-stu-id="a5995-148">Sign in to Union_Lite_2nd using your existing Office 365 administrator credentials.</span></span> <span data-ttu-id="a5995-149">Эти учетные данные будут использоваться для регистрации нового приложения в Azure и настройки параметров, необходимых для единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-149">These credentials will be used to register a new application in Azure and configure the settings required by SSO.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a5995-150">Если на этом этапе для входа в Azure вы используете учетные данные без прав администратора, сценарий `configure-sso` не сможет предоставить согласие администратора для надстройки пользователям в организации.</span><span class="sxs-lookup"><span data-stu-id="a5995-150">If you sign in to Azure using non-administrator credentials during this step, the `configure-sso` script won't be able to provide administrator consent for the add-in to users within your organization.</span></span> <span data-ttu-id="a5995-151">В этом случае единый вход будет недоступен для пользователей надстройки, и им будет предложено выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="a5995-151">SSO will therefore not be available to users of the add-in and they'll be prompted to sign-in.</span></span>

4. <span data-ttu-id="a5995-152">После ввода учетных данных закройте окно браузера и вернитесь к командной строке.</span><span class="sxs-lookup"><span data-stu-id="a5995-152">After you enter your credentials, close the browser window and return to the command prompt.</span></span> <span data-ttu-id="a5995-153">В процессе настройки единого входа на консоль будут выводиться сообщения о состоянии.</span><span class="sxs-lookup"><span data-stu-id="a5995-153">As the SSO configuration process continues, you'll see status messages being written to the console.</span></span> <span data-ttu-id="a5995-154">В соответствии с ними, файлы проекта надстройки, созданные генератором Yeoman, автоматически обновляются с учетом данных, необходимых для процесса единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-154">As described in the console messages, files within the add-in project that the Yeoman generator created are automatically updated with data that's required by the SSO process.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="a5995-155">Попробуйте</span><span class="sxs-lookup"><span data-stu-id="a5995-155">Try it out</span></span>

1. <span data-ttu-id="a5995-156">Когда процесс настройки единого входа будет завершен, для построения проекта, запуска локального веб-сервера и загрузки своей надстройки в ранее выбранное клиентское приложение Office запустите указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="a5995-156">When the SSO configuration process completes, run the following command to build the project, start the local web server, and sideload your add-in in the previously selected Office client application.</span></span>

    > [!NOTE]
    > <span data-ttu-id="a5995-157">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="a5995-157">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="a5995-158">Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="a5995-158">If you are prompted to install a certificate after you run the following command, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    ```command&nbsp;line
    npm start
    ```

2. <span data-ttu-id="a5995-159">Убедитесь, что в клиентском приложении Office (например, Excel, Word или PowerPoint), которое открывается при запуске указанной выше команды, вы выполнили вход как участник той же организации Office 365, что и администратор, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на этапе 3, описанном в [предыдущем разделе](#configure-sso).</span><span class="sxs-lookup"><span data-stu-id="a5995-159">In the Office client application that opens when you run the previous command (i.e., Excel, Word or PowerPoint), make sure that you're signed in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso).</span></span> <span data-ttu-id="a5995-160">Благодаря этому будут созданы соответствующие условия для успешного единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-160">Doing so establishes the appropriate conditions for SSO to succeed.</span></span> 

3. <span data-ttu-id="a5995-161">В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="a5995-161">In Word, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span> <span data-ttu-id="a5995-162">На рисунке ниже показана эта кнопка в Excel. </span><span class="sxs-lookup"><span data-stu-id="a5995-162">The following image shows this button in Excel.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="a5995-164">В нижней части области задач нажмите кнопку **Получить сведения о моем профиле пользователя**, чтобы начать процесс единого входа.</span><span class="sxs-lookup"><span data-stu-id="a5995-164">At the bottom of the task pane, choose the **Get My User Profile Information** button to initiate the SSO process.</span></span> 

    > [!NOTE] 
    > <span data-ttu-id="a5995-165">Если на этом этапе вы еще не вошли в Office, вам будет предложено сделать это.</span><span class="sxs-lookup"><span data-stu-id="a5995-165">If you're not already signed in to Office at this point, you'll be prompted to sign in.</span></span> <span data-ttu-id="a5995-166">Как говорилось выше, вам нужно выполнить вход в качестве участника той же организации Office 365, что и администратор, учетную запись которого вы использовали для подключения к Azure в процессе настройки единого входа на этапе 3, описанном в [предыдущем разделе](#configure-sso). Это необходимо для успешного единого входа. </span><span class="sxs-lookup"><span data-stu-id="a5995-166">As described previously, you should sign in with a user that's a member of the same Office 365 organization as the Office 365 administrator account that you used to connect to Azure while configuring SSO in step 3 of the [previous section](#configure-sso), if you want SSO to succeed.</span></span>

5. <span data-ttu-id="a5995-167">Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя.</span><span class="sxs-lookup"><span data-stu-id="a5995-167">If a dialog window appears to request permissions on behalf of the add-in, this means that SSO is not supported for your scenario and the add-in has instead fallen back to an alternate method of user authentication.</span></span> <span data-ttu-id="a5995-168">Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи).</span><span class="sxs-lookup"><span data-stu-id="a5995-168">This may occur when the tenant administrator hasn't granted consent for the add-in to access Microsoft Graph, or when the user isn't signed into Office with a valid Microsoft Account or Office 365 ("Work or School") account.</span></span> <span data-ttu-id="a5995-169">Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="a5995-169">Choose the **Accept** button in the dialog window to continue.</span></span>

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > <span data-ttu-id="a5995-171">После принятия пользователем запрос разрешений больше не выводится на экран.</span><span class="sxs-lookup"><span data-stu-id="a5995-171">After a user accepts this permissions request, they won't be prompted again in the future.</span></span>

6. <span data-ttu-id="a5995-172">Надстройка получает сведения о профиле пользователя, выполнившего вход, и вносит их в документ.</span><span class="sxs-lookup"><span data-stu-id="a5995-172">The add-in retrieves profile information for the signed-in user and writes it to the document.</span></span> <span data-ttu-id="a5995-173">На приведенном ниже рисунке показан пример сведений о профиле, внесенных на лист Excel.</span><span class="sxs-lookup"><span data-stu-id="a5995-173">The following image shows an example of profile information written to an Excel worksheet.</span></span>

    ![Сведения о профиле пользователя на листе Excel](../images/sso-user-profile-info-excel.png)

## <a name="next-steps"></a><span data-ttu-id="a5995-175">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="a5995-175">Next steps</span></span>

<span data-ttu-id="a5995-176">Поздравляем! Вы успешно создали надстройку области задач, в которой используется единый вход, когда это возможно, и альтернативный метод проверки подлинности пользователей, если единый вход не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="a5995-176">Congratulations, you've successfully created a task pane add-in that uses SSO when possible, and uses an alternate method of user authentication when SSO is not supported.</span></span> <span data-ttu-id="a5995-177">Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).</span><span class="sxs-lookup"><span data-stu-id="a5995-177">To learn more about SSO configuration steps that the Yeoman generator completed automatically, and the code that facilitates the SSO process, see the [Create a Node.js Office Add-in that uses single sign-on](../develop/create-sso-office-add-ins-nodejs.md) tutorial.</span></span>

## <a name="see-also"></a><span data-ttu-id="a5995-178">См. также</span><span class="sxs-lookup"><span data-stu-id="a5995-178">See also</span></span>

- [<span data-ttu-id="a5995-179">Включение единого входа для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a5995-179">Enable single sign-on for Office Add-ins</span></span>](../develop/sso-in-office-add-ins.md)
- [<span data-ttu-id="a5995-180">Создание надстройки Office на платформе Node.js с использованием единого входа</span><span class="sxs-lookup"><span data-stu-id="a5995-180">Create a Node.js Office Add-in that uses single sign-on</span></span>](../develop/create-sso-office-add-ins-nodejs.md)
- [<span data-ttu-id="a5995-181">Устранение ошибок единого входа</span><span class="sxs-lookup"><span data-stu-id="a5995-181">Troubleshoot error messages for single sign-on (SSO)</span></span>](../develop/troubleshoot-sso-in-office-add-ins.md)