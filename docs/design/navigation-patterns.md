---
title: Шаблоны навигации для надстроек Office
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: f0482f7742c6fab97fe563d61d21135c072f8f8f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449147"
---
# <a name="navigation-patterns"></a><span data-ttu-id="a4640-102">Шаблоны навигации</span><span class="sxs-lookup"><span data-stu-id="a4640-102">Navigation patterns</span></span>

<span data-ttu-id="a4640-103">Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную область экрана.</span><span class="sxs-lookup"><span data-stu-id="a4640-103">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="a4640-104">Важно, чтобы навигация была интуитивно понятной, обеспечивала контекст и позволяла пользователю легко перемещаться по всей надстройке.</span><span class="sxs-lookup"><span data-stu-id="a4640-104">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="a4640-105">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="a4640-105">Best practices</span></span>

| <span data-ttu-id="a4640-106">Правильно</span><span class="sxs-lookup"><span data-stu-id="a4640-106">Do</span></span>    | <span data-ttu-id="a4640-107">Неправильно</span><span class="sxs-lookup"><span data-stu-id="a4640-107">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="a4640-108">Убедитесь, что пользователю доступен хорошо видимый параметр навигации.</span><span class="sxs-lookup"><span data-stu-id="a4640-108">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="a4640-109">Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="a4640-109">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="a4640-110">Используйте, по возможности, указанные ниже компоненты, позволяющие пользователям перемещаться по вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="a4640-110">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="a4640-111">Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке</span><span class="sxs-lookup"><span data-stu-id="a4640-111">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="a4640-112">Панель команд</span><span class="sxs-lookup"><span data-stu-id="a4640-112">Command Bar</span></span>

<span data-ttu-id="a4640-113">CommandBar — это область, в которой размещаются команды, работающие с содержимым окна, панели или родительской области, над которой она расположена.</span><span class="sxs-lookup"><span data-stu-id="a4640-113">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="a4640-114">Дополнительные функции включают точку доступа к меню "гамбургер", поиск и боковые команды.</span><span class="sxs-lookup"><span data-stu-id="a4640-114">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Команды: спецификации для области задач рабочего стола](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="a4640-116">Панель вкладок</span><span class="sxs-lookup"><span data-stu-id="a4640-116">Tab Bar</span></span>

<span data-ttu-id="a4640-117">Показывает панель навигации, используя кнопки с расположенными по вертикали текстом и значками.</span><span class="sxs-lookup"><span data-stu-id="a4640-117">Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="a4640-118">Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.</span><span class="sxs-lookup"><span data-stu-id="a4640-118">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Панель вкладок: спецификации для области задач рабочего стола](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="a4640-120">Кнопка "Назад"</span><span class="sxs-lookup"><span data-stu-id="a4640-120">Back Button</span></span>

<span data-ttu-id="a4640-121">Кнопка "Назад" позволяет пользователям восстанавливаться после детализированного навигационного действия.</span><span class="sxs-lookup"><span data-stu-id="a4640-121">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="a4640-122">Этот шаблон помогает пользователям следовать упорядоченной последовательности действий.</span><span class="sxs-lookup"><span data-stu-id="a4640-122">This pattern helps ensure users follow an ordered series of steps.</span></span>  

![Кнопка "Назад": спецификации для области задач рабочего стола](../images/add-in-back-button.png)
