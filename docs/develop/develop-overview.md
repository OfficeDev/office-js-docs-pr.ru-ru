---
title: Разработка надстроек Office
description: Общие сведения о разработке надстроек Office.
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: 419880e8872df20be5a3de40f480f70be2b18859
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292780"
---
# <a name="develop-office-add-ins"></a><span data-ttu-id="bdb3e-103">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-103">Develop Office Add-ins</span></span>

> [!TIP]
> <span data-ttu-id="bdb3e-104">Перед прочтением этой статьи ознакомьтесь со статьей [Создание надстроек Office](../overview/office-add-ins-fundamentals.md).</span><span class="sxs-lookup"><span data-stu-id="bdb3e-104">Please review [Building Office Add-ins](../overview/office-add-ins-fundamentals.md) before reading this article.</span></span>

<span data-ttu-id="bdb3e-105">Все надстройки Office построены на базе платформы надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="bdb3e-105">All Office Add-ins are built upon the Office Add-ins platform.</span></span> <span data-ttu-id="bdb3e-106">Они используют общую структуру, с помощью которой можно реализовать определенные возможности.</span><span class="sxs-lookup"><span data-stu-id="bdb3e-106">They share a common framework through which certain capabilities can be implemented.</span></span> <span data-ttu-id="bdb3e-107">Для каждой создаваемой надстройки следует понять важные принципы, такие как доступность клиентского приложения и платформы, шаблоны программирования API JavaScript для Office, настройку параметров и возможностей надстройки в файле манифеста и т. д.</span><span class="sxs-lookup"><span data-stu-id="bdb3e-107">For any add-in you build, you'll need to understand important concepts like application and platform availability, Office JavaScript API programming patterns, how to specify an add-in's settings and capabilities in the manifest file, and more.</span></span> <span data-ttu-id="bdb3e-108">Эти основные принципы разработки рассматриваются ниже в разделе документации **Основные принципы** > **Разработка**.</span><span class="sxs-lookup"><span data-stu-id="bdb3e-108">Core development concepts like these are covered here in the **Core concepts** > **Develop** section of the documentation.</span></span> <span data-ttu-id="bdb3e-109">Ознакомьтесь с этими сведениями перед изучением документации для клиентского приложения, надстройку для которого вы создаете (например, [Excel](../excel/index.yml)).</span><span class="sxs-lookup"><span data-stu-id="bdb3e-109">Review the information here before exploring the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span>

> [!NOTE]
> <span data-ttu-id="bdb3e-110">Раздел этой документации **Основные понятия** > **Разработка** > **Практическое руководство** включает статьи, посвященные определенным понятиям или задачам разработки.  </span><span class="sxs-lookup"><span data-stu-id="bdb3e-110">The **Core concepts** > **Develop** > **How to** section of this documentation contains articles focused on specific development concepts or tasks.</span></span> <span data-ttu-id="bdb3e-111">Например, здесь можно найти сведения о таких задачах, как [разработка надстроек с Visual Studio Code](develop-add-ins-vscode.md), [автоматическое открытие области задач с документом](automatically-open-a-task-pane-with-a-document.md), [создание команд надстройки](create-addin-commands.md) и [открытие диалогового окна](dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="bdb3e-111">For example, you'll find information there about tasks like [developing add-ins with Visual Studio Code](develop-add-ins-vscode.md), [automatically opening a task pane with a document](automatically-open-a-task-pane-with-a-document.md), [creating add-in commands](create-addin-commands.md), and [opening a dialog box](dialog-api-in-office-add-ins.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="bdb3e-112">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="bdb3e-112">Next steps</span></span>

<span data-ttu-id="bdb3e-113">Ознакомившись с основными понятиями, рассмотренными здесь, изучите документацию для клиентского приложения, надстройку для которого вы создаете (например, [Excel](../excel/index.yml)).</span><span class="sxs-lookup"><span data-stu-id="bdb3e-113">After you're familiar with the core concepts covered here, explore the application-specific documentation that corresponds to the add-in you're building (for example, [Excel](../excel/index.yml)).</span></span> <span data-ttu-id="bdb3e-114">В каждом разделе документации для клиентского приложения содержатся сведения о создании надстроек для определенного клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="bdb3e-114">Each application-specific section of the documentation contains information specifically about building add-ins for a certain Office application.</span></span>

## <a name="see-also"></a><span data-ttu-id="bdb3e-115">См. также</span><span class="sxs-lookup"><span data-stu-id="bdb3e-115">See also</span></span>

- [<span data-ttu-id="bdb3e-116">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-116">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="bdb3e-117">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-117">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="bdb3e-118">Основные принципы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-118">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="bdb3e-119">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-119">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="bdb3e-120">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-120">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="bdb3e-121">Публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="bdb3e-121">Publish Office Add-ins</span></span>](../publish/publish.md)
