---
title: Отладка надстроек Office на Mac
description: ''
ms.date: 11/26/2019
localization_priority: Priority
ms.openlocfilehash: d7609ef3952a68afc4db18100e24b3c1b3947dac
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217325"
---
# <a name="debug-office-add-ins-on-a-mac"></a><span data-ttu-id="5e119-102">Отладка надстроек Office на Mac</span><span class="sxs-lookup"><span data-stu-id="5e119-102">Debug Office Add-ins on a Mac</span></span>

<span data-ttu-id="5e119-p101">Так как надстройки создаются с помощью кода HTML и JavaScript, они рассчитаны на работу на многих платформах, но отрисовка HTML в разных браузерах может слегка отличаться. В этой статье описывается отладка надстроек на компьютере Mac.</span><span class="sxs-lookup"><span data-stu-id="5e119-p101">Because add-ins are developed using HTML and JavaScript, they are designed to work across platforms, but there might be subtle differences in how different browsers render the HTML. This article describes how to debug add-ins running on a Mac.</span></span>

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a><span data-ttu-id="5e119-105">Отладка с помощью Safari Web Inspector на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="5e119-105">Debugging with Safari Web Inspector on a Mac</span></span>

<span data-ttu-id="5e119-106">Если у вас есть надстройка, которая отображает пользовательский интерфейс в области задач или контентной надстройке, вы можете отлаживать надстройку Office с помощью Safari Web Inspector.</span><span class="sxs-lookup"><span data-stu-id="5e119-106">If you have add-in that shows UI in a task pane or in a content add-in, you can debug an Office Add-in using Safari Web Inspector.</span></span>

<span data-ttu-id="5e119-107">Отладку надстроек Office на компьютере Mac можно выполнить, только если на нем установлена система Mac OS High Sierra И Office для Mac версии 16.9.1 (сборка 18012504) или более поздней.</span><span class="sxs-lookup"><span data-stu-id="5e119-107">To be able to debug Office Add-ins on Mac, you must have Mac OS High Sierra AND Mac Office Version: 16.9.1 (Build 18012504) or later.</span></span> <span data-ttu-id="5e119-108">Если у вас нет сборки Office для Mac, вы можете получить ее, присоединившись к [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="5e119-108">If you don't have an Office Mac build, you can get one by joining the [Office 365 Developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="5e119-109">Для этого откройте терминал и установите свойство `OfficeWebAddinDeveloperExtras` для соответствующего приложения Office следующим образом:</span><span class="sxs-lookup"><span data-stu-id="5e119-109">To start, open a terminal and set the `OfficeWebAddinDeveloperExtras` property for the relevant Office application as follows:</span></span>

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

<span data-ttu-id="5e119-110">Затем откройте приложение Office и [загрузите свою неопубликованную надстройку](sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="5e119-110">Then, open the Office application and [sideload your add-in](sideload-an-office-add-in-on-ipad-and-mac.md).</span></span> <span data-ttu-id="5e119-111">Щелкните надстройку правой кнопкой мыши. В контекстном меню отобразится пункт **Проверить элемент**.</span><span class="sxs-lookup"><span data-stu-id="5e119-111">Right-click the add-in and you should see an **Inspect Element** option in the context menu.</span></span> <span data-ttu-id="5e119-112">Выберите его. Он появится в инспекторе, где можно устанавливать точки останова и отлаживать надстройку.</span><span class="sxs-lookup"><span data-stu-id="5e119-112">Select that option and it will pop the Inspector, where you can set breakpoints and debug your add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="5e119-113">Если при попытке использовать инспектор диалоговое окно мерцает, обновите Office до последней версии.</span><span class="sxs-lookup"><span data-stu-id="5e119-113">If you're trying to use the inspector and the dialog flickers, update Office to the latest version.</span></span> <span data-ttu-id="5e119-114">Если проблема с мерцанием сохраняется, попробуйте применить следующее временное решение:</span><span class="sxs-lookup"><span data-stu-id="5e119-114">If that doesn't resolve the flickering, try the following workaround:</span></span>
> 1. <span data-ttu-id="5e119-115">Уменьшите размер диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="5e119-115">Reduce the size of the dialog.</span></span>
> 2. <span data-ttu-id="5e119-116">Выберите пункт **Проверить элемент**, который откроется в новом окне.</span><span class="sxs-lookup"><span data-stu-id="5e119-116">Choose **Inspect Element**, which opens in a new window.</span></span>
> 3. <span data-ttu-id="5e119-117">Измените размер диалогового окна на исходный.</span><span class="sxs-lookup"><span data-stu-id="5e119-117">Resize the dialog to its original size.</span></span>
> 4. <span data-ttu-id="5e119-118">Используйте инспектор должным образом.</span><span class="sxs-lookup"><span data-stu-id="5e119-118">Use the inspector as required.</span></span>

## <a name="clearing-the-office-applications-cache-on-a-mac"></a><span data-ttu-id="5e119-119">Очистка кэша приложения Office на компьютере Mac</span><span class="sxs-lookup"><span data-stu-id="5e119-119">Clearing the Office application's cache on a Mac</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]
