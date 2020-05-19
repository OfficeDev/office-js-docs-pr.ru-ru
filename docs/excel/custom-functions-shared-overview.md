---
ms.date: 05/17/2020
description: Узнайте, как запускать пользовательские функции, кнопки ленты и код области задач в одной и той же среде выполнения JavaScript для координации сценариев в вашей надстройке.
title: Выполнение кода надстройки в общей среде выполнения JavaScript
localization_priority: Priority
ms.openlocfilehash: afb07c5223e26ba1e1adbf40c7a4b2e4f7c06349
ms.sourcegitcommit: 54e2892c0c26b9ad1e4dba8aba48fea39f853b6c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44275933"
---
# <a name="overview-run-your-add-in-code-in-a-shared-javascript-runtimes"></a><span data-ttu-id="3ad13-103">Обзор: выполнение кода надстройки в общедоступных средах выполнения JavaScript</span><span class="sxs-lookup"><span data-stu-id="3ad13-103">Overview: Run your add-in code in a shared JavaScript runtimes</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="3ad13-104">При запуске Excel на компьютере с Windows или на Mac надстройка запустит код для кнопок ленты, пользовательских функций и области задач в отдельных средах выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3ad13-104">When running Excel on Windows or Mac, your add-in will run code for ribbon buttons, custom functions, and the task pane in separate JavaScript runtime environments.</span></span> <span data-ttu-id="3ad13-105">Из-за этого возникают ограничения, например невозможность удобно предоставлять общий доступ к глобальным данным и отсутствие доступа ко всей функциональности CORS для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="3ad13-105">This creates limitations such as not being able to easily share global data, and not being able to access all CORS functionality from a custom function.</span></span>

<span data-ttu-id="3ad13-106">Тем не менее, можно настроить надстройку Excel так, чтобы обеспечить общий доступ к коду в одной и той же среде выполнения JavaScript (то есть в общей среде выполнения).</span><span class="sxs-lookup"><span data-stu-id="3ad13-106">However, you can configure your Excel add-in to share code in the same JavaScript runtime (also referred to as a shared runtime).</span></span> <span data-ttu-id="3ad13-107">За счет этого повышается скоординированность работы надстройки и упрощается доступ к DOM панели задач и CORS из всех компонентов надстройки.</span><span class="sxs-lookup"><span data-stu-id="3ad13-107">This enables better coordination across your add-in and access to the task pane DOM and CORS from all parts of your add-in.</span></span>

<span data-ttu-id="3ad13-108">При настройке общей среды выполнения становятся возможными следующие сценарии:</span><span class="sxs-lookup"><span data-stu-id="3ad13-108">Configuring a shared runtime enables the following scenarios:</span></span>

- <span data-ttu-id="3ad13-109">У вашей настройки будет общая модель DOM, доступная для ленты, области задач и пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="3ad13-109">Your add-in will have a shared DOM that the ribbon, task pane, and custom functions can all access.</span></span>
- <span data-ttu-id="3ad13-110">У ваших пользовательских функций будет полная поддержка CORS.</span><span class="sxs-lookup"><span data-stu-id="3ad13-110">Your custom functions will have full CORS support.</span></span>
- <span data-ttu-id="3ad13-111">Пользовательские функции могут вызывать API Office.js для чтения данных из электронной таблицы.</span><span class="sxs-lookup"><span data-stu-id="3ad13-111">Your custom functions can call Office.js APIs to read spreadsheet document data.</span></span>
- <span data-ttu-id="3ad13-112">Надстройка может выполнять код сразу после открытия документа.</span><span class="sxs-lookup"><span data-stu-id="3ad13-112">Your add-in can run code as soon as the document is opened.</span></span>
- <span data-ttu-id="3ad13-113">Надстройка может продолжать выполнять код после закрытия области задач.</span><span class="sxs-lookup"><span data-stu-id="3ad13-113">Your add-in can continue running code after the task pane is closed.</span></span>

<span data-ttu-id="3ad13-114">При выполнении пользовательских функций в общей среде выполнения с областью задач эта среда будет выполняться в экземпляре браузера на различных платформах, как описано в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md). Кроме того, все кнопки, отображаемые вашей надстройкой Excel на ленте, будут выполняться в этой же общей среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="3ad13-114">When you run custom functions in a shared runtime with the task pane, it will run in a browser instance on different platforms as explained in [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md). Additionally, any buttons that your Excel add-in displays on the ribbon will run in the same shared runtime.</span></span> <span data-ttu-id="3ad13-115">На следующем рисунке показано, как пользовательские функции, пользовательский интерфейс ленты и код области задач будут запускаться в одной и той же среде выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3ad13-115">The following image shows how custom functions, the ribbon UI, and the task pane code will all run in the same JavaScript runtime.</span></span>

![Пользовательские функции, выполняемые в общедоступной среде выполнения с кнопками ленты и областью задач в Excel](../images/custom-functions-in-browser-runtime.png)

## <a name="set-up-a-shared-runtime"></a><span data-ttu-id="3ad13-117">Настройка общей среды выполнения</span><span class="sxs-lookup"><span data-stu-id="3ad13-117">Set up a shared runtime</span></span>

<span data-ttu-id="3ad13-118">В [статье Настройка общедоступной среды выполнения](./configure-your-add-in-to-use-a-shared-runtime.md) , в которой описано, как настроить пользовательские функции для использования общей среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="3ad13-118">See the [configuring a shared runtime article](./configure-your-add-in-to-use-a-shared-runtime.md) to learn how to set up your custom functions to use a shared runtime.</span></span>

### <a name="debugging"></a><span data-ttu-id="3ad13-119">Отладка</span><span class="sxs-lookup"><span data-stu-id="3ad13-119">Debugging</span></span>

<span data-ttu-id="3ad13-120">В настоящее время при использовании общей среды выполнения невозможно использовать Visual Studio Code для отладки пользовательских функций в Excel под управлением Windows.</span><span class="sxs-lookup"><span data-stu-id="3ad13-120">When using a shared runtime, you can't use Visual Studio Code to debug custom functions in Excel on Windows at this time.</span></span> <span data-ttu-id="3ad13-121">Вместо этого вам потребуется использовать инструменты разработчика.</span><span class="sxs-lookup"><span data-stu-id="3ad13-121">You'll need to use developer tools instead.</span></span> <span data-ttu-id="3ad13-122">Дополнительные сведения см. в статье [Отладка надстроек с помощью средств разработчика в Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span><span class="sxs-lookup"><span data-stu-id="3ad13-122">For more information, see [Debug add-ins using developer tools on Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).</span></span>

## <a name="give-us-feedback"></a><span data-ttu-id="3ad13-123">Отправьте нам отзыв</span><span class="sxs-lookup"><span data-stu-id="3ad13-123">Give us feedback</span></span>

<span data-ttu-id="3ad13-124">Мы будем рады услышать ваши отзывы об этой функции.</span><span class="sxs-lookup"><span data-stu-id="3ad13-124">We'd love to hear your feedback on this feature.</span></span> <span data-ttu-id="3ad13-125">Если вы обнаружите какие-либо ошибки или проблемы, если у вас есть запросы относительно этой функции, сообщите нам, создав проблему GitHub в [репозитории office-js](https://github.com/OfficeDev/office-js).</span><span class="sxs-lookup"><span data-stu-id="3ad13-125">If you find any bugs, issues, or have requests on this feature, please let us know by creating a GitHub issue in the [office-js repo](https://github.com/OfficeDev/office-js).</span></span>

## <a name="see-also"></a><span data-ttu-id="3ad13-126">См. также</span><span class="sxs-lookup"><span data-stu-id="3ad13-126">See also</span></span>

- [<span data-ttu-id="3ad13-127">Руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач</span><span class="sxs-lookup"><span data-stu-id="3ad13-127">Tutorial: Share data and events between Excel custom functions and the task pane</span></span>](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [<span data-ttu-id="3ad13-128">Вызов API Excel из пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="3ad13-128">Call Excel APIs from your custom function</span></span>](call-excel-apis-from-custom-function.md)
