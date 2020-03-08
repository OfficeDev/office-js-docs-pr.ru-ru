### <a name="configuration"></a><span data-ttu-id="afc5a-101">Конфигурация</span><span class="sxs-lookup"><span data-stu-id="afc5a-101">Configuration</span></span>

<span data-ttu-id="afc5a-102">В следующих файлах указываются параметры конфигурации для надстройки.</span><span class="sxs-lookup"><span data-stu-id="afc5a-102">The following files specify configuration settings for the add-in.</span></span>

- <span data-ttu-id="afc5a-103">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="afc5a-103">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>

- <span data-ttu-id="afc5a-104">**./. ENV** в корневом каталоге проекта определяет константы, используемые в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="afc5a-104">The **./.ENV** file in the root directory of the project defines constants that are used by the add-in project.</span></span>

### <a name="task-pane"></a><span data-ttu-id="afc5a-105">Для области задач</span><span class="sxs-lookup"><span data-stu-id="afc5a-105">Task pane</span></span> 

<span data-ttu-id="afc5a-106">Следующие файлы определяют пользовательский интерфейс и функциональные возможности области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="afc5a-106">The following files define the add-in's task pane UI and functionality.</span></span>

- <span data-ttu-id="afc5a-107">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="afc5a-107">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>

- <span data-ttu-id="afc5a-108">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="afc5a-108">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>

- <span data-ttu-id="afc5a-109">В проекте JavaScript файл **./СРК/таскпане/таскпане.ЖС** содержит код для инициализации надстройки.</span><span class="sxs-lookup"><span data-stu-id="afc5a-109">In a JavaScript project, the **./src/taskpane/taskpane.js** file contains code to initialize the add-in.</span></span> <span data-ttu-id="afc5a-110">В проекте TypeScript файл **./СРК/таскпане/таскпане.ТС** содержит код для инициализации надстройки, а также код, использующий библиотеку JavaScript для Office для добавления данных из Microsoft Graph в документ Office.</span><span class="sxs-lookup"><span data-stu-id="afc5a-110">In a TypeScript project, the **./src/taskpane/taskpane.ts** file contains code to initialize the add-in and also code that uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span>

### <a name="authentication"></a><span data-ttu-id="afc5a-111">Проверка подлинности</span><span class="sxs-lookup"><span data-stu-id="afc5a-111">Authentication</span></span>

<span data-ttu-id="afc5a-112">Следующие файлы упрощают процесс единого входа и записывают данные в документ Office.</span><span class="sxs-lookup"><span data-stu-id="afc5a-112">The following files facilitate the SSO process and write data to the Office document.</span></span>

- <span data-ttu-id="afc5a-113">В проекте JavaScript файл **./СРК/Хелперс/докуменселпер.ЖС** содержит код, который использует библиотеку JavaScript для Office для добавления данных из Microsoft Graph в документ Office.</span><span class="sxs-lookup"><span data-stu-id="afc5a-113">In a JavaScript project, the **./src/helpers/documentHelper.js** file contains code that uses the Office JavaScript library to add the data from Microsoft Graph to the Office document.</span></span> <span data-ttu-id="afc5a-114">В проекте TypeScript нет такого файла; код, использующий библиотеку JavaScript для Office для добавления данных из Microsoft Graph в документ Office, существует в файле **./СРК/таскпане/таскпане.ТС** .</span><span class="sxs-lookup"><span data-stu-id="afc5a-114">There is no such file in a TypeScript project; the code that uses the Office JavaScript library to add the data from Microsoft Graph to the Office document exists in **./src/taskpane/taskpane.ts** instead.</span></span>

- <span data-ttu-id="afc5a-115">Файл **./src/helpers/fallbackauthdialog.HTML** — это страница без пользовательского интерфейса, которая загружает JavaScript для стратегии резервной проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="afc5a-115">The **./src/helpers/fallbackauthdialog.html** file is the UI-less page that loads the JavaScript for the fallback authentication strategy.</span></span>

- <span data-ttu-id="afc5a-116">Файл **./СРК/Хелперс/фаллбаккаусдиалог.ЖС** содержит JavaScript для стратегии проверки подлинности, которая подписывает пользователя с помощью msal. js.</span><span class="sxs-lookup"><span data-stu-id="afc5a-116">The **./src/helpers/fallbackauthdialog.js** file contains the JavaScript for the fallback authentication strategy that signs in the user with msal.js.</span></span>

- <span data-ttu-id="afc5a-117">Файл **./СРК/Хелперс/фаллбаккаусхелпер.ЖС** содержит область задач JavaScript, которая вызывает резервную стратегию проверки подлинности в сценариях, когда проверка подлинности SSO не поддерживается.</span><span class="sxs-lookup"><span data-stu-id="afc5a-117">The **./src/helpers/fallbackauthhelper.js** file contains the task pane JavaScript that invokes the fallback authentication strategy in scenarios when SSO authentication is not supported.</span></span>

- <span data-ttu-id="afc5a-118">Файл **./src/helpers/ssoauthhelper.js** содержит вызов JavaScript для API единого входа, `getAccessToken`, получает маркер начальной загрузки, инициирует его замену на маркер доступа для Microsoft Graph и вызывает данные Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="afc5a-118">The **./src/helpers/ssoauthhelper.js** file contains the JavaScript call to the SSO API, `getAccessToken`, receives the bootstrap token, initiates the swap of the bootstrap token for an access token to Microsoft Graph, and calls to Microsoft Graph for the data.</span></span>