---
title: Шаблоны навигации для надстроек Office
description: ''
ms.date: 06/26/2018
ms.openlocfilehash: b7fee6fad703ce7c8f4c5f8b848d6bf28b239b09
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432972"
---
# <a name="navigation-patterns"></a><span data-ttu-id="ea8f1-102">Шаблоны навигации</span><span class="sxs-lookup"><span data-stu-id="ea8f1-102">Navigation patterns</span></span>

<span data-ttu-id="ea8f1-103">Доступ к основным функциям надстройки осуществляется через определенные типы команд и ограниченную область экрана.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-103">The main features of an add-in are accessed through specific command types and limited screen area.</span></span> <span data-ttu-id="ea8f1-104">Важно, чтобы навигация была интуитивно понятной, обеспечивала контекст и позволяла пользователю легко перемещаться по всей надстройке.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-104">It is important that navigation is intuitive, provides context, and allows the user to move easily throughout the add-in.</span></span>

## <a name="best-practices"></a><span data-ttu-id="ea8f1-105">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="ea8f1-105">Best practices</span></span>

| <span data-ttu-id="ea8f1-106">Правильно</span><span class="sxs-lookup"><span data-stu-id="ea8f1-106">Do</span></span>    | <span data-ttu-id="ea8f1-107">Неправильно</span><span class="sxs-lookup"><span data-stu-id="ea8f1-107">Don't</span></span> |
| :---- | :---- |
| <span data-ttu-id="ea8f1-108">Убедитесь, что пользователю доступен хорошо видимый параметр навигации.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-108">Ensure the user has a clearly visible navigation option.</span></span> | <span data-ttu-id="ea8f1-109">Не затрудняйте процесс навигации, используя нестандартный пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-109">Don't complicate the navigation process by using non-standard UI.</span></span>
| <span data-ttu-id="ea8f1-110">Используйте, по возможности, указанные ниже компоненты, позволяющие пользователям перемещаться по вашей надстройке.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-110">Utilize the following components as applicable to allow users to navigate through your add-in.</span></span> | <span data-ttu-id="ea8f1-111">Не затрудняйте понимание пользователем своего текущего места или контекста в надстройке</span><span class="sxs-lookup"><span data-stu-id="ea8f1-111">Don't make it difficult for the user to understand their current place or context within the add-in</span></span>



## <a name="command-bar"></a><span data-ttu-id="ea8f1-112">Панель команд</span><span class="sxs-lookup"><span data-stu-id="ea8f1-112">UserForm Command Bar</span></span>

<span data-ttu-id="ea8f1-113">CommandBar — это область, в которой размещаются команды, работающие с содержимым окна, панели или родительской области, над которой она расположена.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-113">CommandBar is a surface that houses commands that operate on the content of the window, panel, or parent region it resides above.</span></span> <span data-ttu-id="ea8f1-114">Дополнительные функции включают точку доступа к меню "гамбургер", поиск и боковые команды.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-114">Optional features include a hamburger menu access point, search, and side commands.</span></span>

![Команды: спецификации для области задач рабочего стола](../images/add-in-command-bar.png)



## <a name="tab-bar"></a><span data-ttu-id="ea8f1-116">Панель вкладок</span><span class="sxs-lookup"><span data-stu-id="ea8f1-116">Tab bar</span></span>

<span data-ttu-id="ea8f1-117">Показывает панель навигации, используя кнопки с расположенными по вертикали текстом и значками.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-117">Tab bar - Shows navigation using buttons with vertically stacked text and icons.</span></span> <span data-ttu-id="ea8f1-118">Панель вкладок обеспечивает навигацию с помощью вкладок с короткими и понятными названиями.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-118">Use the tab bar to provide navigation using tabs with short and descriptive titles.</span></span>

![Панель вкладок: спецификации для области задач рабочего стола](../images/add-in-tab-bar.png)


## <a name="back-button"></a><span data-ttu-id="ea8f1-120">Кнопка "Назад"</span><span class="sxs-lookup"><span data-stu-id="ea8f1-120">Back button</span></span>

<span data-ttu-id="ea8f1-121">Кнопка "Назад" позволяет пользователям восстанавливаться после детализированного навигационного действия.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-121">The back button allows users to recover from a drill down navigational action.</span></span> <span data-ttu-id="ea8f1-122">Этот шаблон помогает пользователям следовать упорядоченной последовательности действий.</span><span class="sxs-lookup"><span data-stu-id="ea8f1-122">Use this pattern to ensure users follow an ordered series of steps.</span></span>  

![Кнопка "Назад": спецификации для области задач рабочего стола](../images/add-in-back-button.png)
