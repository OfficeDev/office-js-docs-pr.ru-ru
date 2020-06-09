---
title: Шаблоны навигации для надстроек Office
description: Ознакомьтесь с рекомендациями по использованию панелей команд, вкладок и кнопок "назад", чтобы разработать навигацию для надстройки Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 812b56edc0653812c3519735a7300e5f3d7b38a6
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608511"
---
# <a name="navigation-patterns"></a><span data-ttu-id="dd64b-103">Шаблоны навигации</span><span class="sxs-lookup"><span data-stu-id="dd64b-103">Navigation patterns</span></span>

<span data-ttu-id="dd64b-104">Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную область экрана.</span><span class="sxs-lookup"><span data-stu-id="dd64b-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="dd64b-105">Важно, чтобы навигация была интуитивно понятной, обеспечивала контекст и позволяла пользователю легко перемещаться по всей надстройке.</span><span class="sxs-lookup"><span data-stu-id="dd64b-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="dd64b-106">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="dd64b-106">Best practices</span></span>

| <span data-ttu-id="dd64b-107">Правильно</span><span class="sxs-lookup"><span data-stu-id="dd64b-107">Do</span></span>    | <span data-ttu-id="dd64b-108">Неправильно</span><span class="sxs-lookup"><span data-stu-id="dd64b-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="dd64b-109">Убедитесь, что пользователю доступен хорошо видимый параметр навигации.</span><span class="sxs-lookup"><span data-stu-id="dd64b-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="dd64b-110">Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="dd64b-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="dd64b-111">Используйте, по возможности, указанные ниже компоненты, позволяющие пользователям перемещаться по вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="dd64b-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="dd64b-112">Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке</span><span class="sxs-lookup"><span data-stu-id="dd64b-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="dd64b-113">Панель команд</span><span class="sxs-lookup"><span data-stu-id="dd64b-113">Command Bar</span></span>

<span data-ttu-id="dd64b-114">CommandBar — это область, в которой размещаются команды, работающие с содержимым окна, панели или родительской области, над которой она расположена.</span><span class="sxs-lookup"><span data-stu-id="dd64b-114">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="dd64b-115">Дополнительные функции включают точку доступа к меню "гамбургер", поиск и боковые команды.</span><span class="sxs-lookup"><span data-stu-id="dd64b-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Команды: спецификации для области задач рабочего стола](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="dd64b-117">Панель вкладок</span><span class="sxs-lookup"><span data-stu-id="dd64b-117">Tab Bar</span></span>

<span data-ttu-id="dd64b-118">Показывает панель навигации, используя кнопки с расположенными по вертикали текстом и значками.</span><span class="sxs-lookup"><span data-stu-id="dd64b-118">Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="dd64b-119">Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.</span><span class="sxs-lookup"><span data-stu-id="dd64b-119">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Панель вкладок: спецификации для области задач рабочего стола](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="dd64b-121">Кнопка "Назад"</span><span class="sxs-lookup"><span data-stu-id="dd64b-121">Back Button</span></span>

<span data-ttu-id="dd64b-122">Кнопка "Назад" позволяет пользователям восстанавливаться после детализированного навигационного действия.</span><span class="sxs-lookup"><span data-stu-id="dd64b-122">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="dd64b-123">Этот шаблон помогает пользователям следовать упорядоченной последовательности действий.</span><span class="sxs-lookup"><span data-stu-id="dd64b-123">This pattern helps ensure users follow an ordered series of steps.</span></span>  

![Кнопка "Назад": спецификации для области задач рабочего стола](../images/add-in-back-button.png)
