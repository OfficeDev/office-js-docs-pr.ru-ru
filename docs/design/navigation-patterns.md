---
title: Шаблоны навигации для надстроек Office
description: Ознакомьтесь с рекомендациями по использованию панелей команд, вкладок и кнопок "назад", чтобы разработать навигацию для надстройки Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 3bb350ede78bef684899f26e4818eba440677541
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132034"
---
# <a name="navigation-patterns"></a><span data-ttu-id="17500-103">Шаблоны навигации</span><span class="sxs-lookup"><span data-stu-id="17500-103">Navigation patterns</span></span>

<span data-ttu-id="17500-104">Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную область экрана.</span><span class="sxs-lookup"><span data-stu-id="17500-104">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="17500-105">Важно, чтобы навигация была интуитивно понятной, обеспечивала контекст и позволяла пользователю легко перемещаться по всей надстройке.</span><span class="sxs-lookup"><span data-stu-id="17500-105">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="17500-106">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="17500-106">Best practices</span></span>

| <span data-ttu-id="17500-107">Правильно</span><span class="sxs-lookup"><span data-stu-id="17500-107">Do</span></span>    | <span data-ttu-id="17500-108">Неправильно</span><span class="sxs-lookup"><span data-stu-id="17500-108">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="17500-109">Убедитесь, что пользователю доступен хорошо видимый параметр навигации.</span><span class="sxs-lookup"><span data-stu-id="17500-109">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="17500-110">Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="17500-110">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="17500-111">Используйте, по возможности, указанные ниже компоненты, позволяющие пользователям перемещаться по вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="17500-111">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="17500-112">Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке</span><span class="sxs-lookup"><span data-stu-id="17500-112">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>

## <a name="command-bar"></a><span data-ttu-id="17500-113">Панель команд</span><span class="sxs-lookup"><span data-stu-id="17500-113">Command Bar</span></span>

<span data-ttu-id="17500-114">Панель элементов управления — это поверхность области задач, в которой размещаются команды, работающие с содержимым окна, панели или родительской области, расположенной выше.</span><span class="sxs-lookup"><span data-stu-id="17500-114">The CommandBar is a surface within the task pane that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="17500-115">Дополнительные функции включают точку доступа к меню "гамбургер", поиск и боковые команды.</span><span class="sxs-lookup"><span data-stu-id="17500-115">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Иллюстрация, демонстрирующая панель команд в области задач приложения Office для настольных ПК.](../images/add-in-command-bar.png)

## <a name="tab-bar"></a><span data-ttu-id="17500-118">Панель вкладок</span><span class="sxs-lookup"><span data-stu-id="17500-118">Tab Bar</span></span>

<span data-ttu-id="17500-119">Панель вкладок показывает навигацию с помощью кнопок с вертикальным текстом и значками.</span><span class="sxs-lookup"><span data-stu-id="17500-119">The tab bar shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="17500-120">Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.</span><span class="sxs-lookup"><span data-stu-id="17500-120">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Иллюстрация, на которой показана панель вкладок в области задач приложения Office для настольных ПК.](../images/add-in-tab-bar.png)

## <a name="back-button"></a><span data-ttu-id="17500-123">Кнопка "Назад"</span><span class="sxs-lookup"><span data-stu-id="17500-123">Back Button</span></span>

<span data-ttu-id="17500-124">Кнопка "назад" позволяет пользователям восстанавливаться при переходе по навигации.</span><span class="sxs-lookup"><span data-stu-id="17500-124">The back button allows users to recover from a drill-down navigational action.</span></span> <span data-ttu-id="17500-125">Этот шаблон помогает пользователям следовать упорядоченной последовательности действий.</span><span class="sxs-lookup"><span data-stu-id="17500-125">This pattern helps ensure users follow an ordered series of steps.</span></span>

![Иллюстрация, на которой показана кнопка "назад" в области задач приложения Office для настольных ПК.](../images/add-in-back-button.png)
