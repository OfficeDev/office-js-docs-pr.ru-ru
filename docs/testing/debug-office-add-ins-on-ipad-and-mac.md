---
title: Отладка надстроек Office на Mac
description: Узнайте, как использовать Mac для отлаки надстроек Office.
ms.date: 10/16/2020
localization_priority: Normal
ms.openlocfilehash: b2164e3ed672b2911db6841fad24441b67882204
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237947"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="35b94-103">Отладка надстроек Office на Mac</span><span class="sxs-lookup"><span data-stu-id="35b94-103">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="35b94-p101">Так как надстройки создаются с помощью кода HTML и JavaScript, они рассчитаны на работу на многих платформах, но отрисовка HTML в разных браузерах может слегка отличаться. В этой статье описывается отладка надстроек на компьютере Mac.</span><span class="sxs-lookup"><span data-stu-id="35b94-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="35b94-106">Отладка с помощью Safari Web Inspector на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="35b94-106">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="35b94-107">Если у вас есть надстройка, которая отображает пользовательский интерфейс в области задач или контентной надстройке, вы можете отлаживать надстройку Office с помощью Safari Web Inspector.</span><span class="sxs-lookup"><span data-stu-id="35b94-107">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="35b94-108">Для отлаки надстроек Office на Mac необходимо иметь Mac OS High Sierra и Mac Office версии 16.9.1 (сборка 18012504) или более поздней.</span><span class="sxs-lookup"><span data-stu-id="35b94-108">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office version 16.9.1 (build 18012504) or later.</span></span> <span data-ttu-id="35b94-109">Если у вас нет сборки Office для Mac, вы можете получить ее, присоединившись к программе для разработчиков [Microsoft 365.](https://developer.microsoft.com/office/dev-program)</span><span class="sxs-lookup"><span data-stu-id="35b94-109">If you don't have an Office Mac build, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="35b94-110">Для этого откройте терминал и установите свойство `OfficeWebAddinDeveloperExtras` для соответствующего приложения Office следующим образом:</span><span class="sxs-lookup"><span data-stu-id="35b94-110">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

    > [!IMPORTANT]
    > <span data-ttu-id="35b94-111">Сборки Office в Магазине приложений Mac не поддерживают `OfficeWebAddinDeveloperExtras` этот флаг.</span><span class="sxs-lookup"><span data-stu-id="35b94-111">Mac App Store builds of Office do not support the `OfficeWebAddinDeveloperExtras` flag.</span></span>

<span data-ttu-id="35b94-112">Затем откройте приложение Office и [загрузите свою неопубликованную надстройку](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="35b94-112">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="35b94-113">Щелкните надстройку правой кнопкой мыши. В контекстном меню отобразится пункт **Проверить элемент**.</span><span class="sxs-lookup"><span data-stu-id="35b94-113">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="35b94-114">Выберите его. Он появится в инспекторе, где можно устанавливать точки останова и отлаживать надстройку.</span><span class="sxs-lookup"><span data-stu-id="35b94-114">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="35b94-115">Если при попытке использовать инспектор диалоговое окно мерцает, обновите Office до последней версии.</span><span class="sxs-lookup"><span data-stu-id="35b94-115">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="35b94-116">Если проблема с мерцанием сохраняется, попробуйте применить следующее временное решение:</span><span class="sxs-lookup"><span data-stu-id="35b94-116">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="35b94-117">Уменьшите размер диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="35b94-117">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="35b94-118">Выберите пункт **Проверить элемент**, который откроется в новом окне.</span><span class="sxs-lookup"><span data-stu-id="35b94-118">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="35b94-119">Измените размер диалогового окна на исходный.</span><span class="sxs-lookup"><span data-stu-id="35b94-119">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="35b94-120">Используйте инспектор должным образом.</span><span class="sxs-lookup"><span data-stu-id="35b94-120">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="35b94-121">Очистка кэша приложения Office на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="35b94-121">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
