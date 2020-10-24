---
title: Разработка надстроек Office с помощью Visual Studio Code
description: Как разрабатывать надстройки Office с помощью Visual Studio Code.
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: 7bd12f6d6e4aff6e8a80f9e9c2e5042b726eca0c
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741101"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="1fcf8-103">Разработка надстроек Office с помощью Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="1fcf8-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="1fcf8-104">В этой статье описано, как разработать надстройку Office с помощью [Visual Studio Code (VS Code)](https://code.visualstudio.com).</span><span class="sxs-lookup"><span data-stu-id="1fcf8-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1fcf8-105">Сведения об использовании Visual Studio для создания надстроек Office см. в статье [Разработка надстроек Office в Visual Studio](develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="1fcf8-105">For information about using Visual Studio to create an Office Add-in, see [Develop Office Add-ins with Visual Studio](develop-add-ins-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="1fcf8-106">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="1fcf8-106">Prerequisites</span></span>

- [<span data-ttu-id="1fcf8-107">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="1fcf8-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="1fcf8-108">Создание проекта надстройки с помощью генератора Yeoman</span><span class="sxs-lookup"><span data-stu-id="1fcf8-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="1fcf8-109">Если вы используете VS Code в качестве интегрированной среды разработки (IDE), следует создать проект надстройки Office с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Генератор Yeoman создает проект Node.js, которым можно управлять с помощью VS Code или любого другого редактора.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="1fcf8-110">Чтобы создать надстройку Office с помощью генератора Yeoman, следуйте указаниям из [5-минутного краткого руководства](/office/dev/add-ins/), соответствующего типу надстройки, которую нужно создать.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](/office/dev/add-ins/) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="1fcf8-111">Разработка надстройки с помощью VS Code</span><span class="sxs-lookup"><span data-stu-id="1fcf8-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="1fcf8-112">Когда генератор Yeoman закончит создание проекта надстройки, откройте корневую папку проекта с помощью VS Code.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="1fcf8-113">В Windows вы можете перейти в корневой каталог проекта с помощью командной строки и ввести `code .`, чтобы открыть эту папку в VS Code.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="1fcf8-114">На компьютере Mac потребуется [добавить в путь команду `code`](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) перед использованием этой команды для открытия папки проекта в VS Code.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="1fcf8-115">Генератор Yeoman создает простую надстройку с ограниченными возможностями.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="1fcf8-116">Вы можете настроить надстройку, изменив файлы [манифеста](add-in-manifests.md), HTML, JavaScript, TypeScript или CSS в VS Code.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="1fcf8-117">Общее описание структуры проекта и файлов в проекте надстройки, созданном генератором Yeoman, см. в рекомендациях по генератору Yeoman в [5-минутном кратком руководстве](/office/dev/add-ins/), соответствующем типу созданной надстройки.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](/office/dev/add-ins/) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="1fcf8-118">Тестирование и отладка надстройки</span><span class="sxs-lookup"><span data-stu-id="1fcf8-118">Test and debug the add-in</span></span>

<span data-ttu-id="1fcf8-119">Методы тестирования, отладки и устранения неполадок надстроек Office зависят от платформы.</span><span class="sxs-lookup"><span data-stu-id="1fcf8-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="1fcf8-120">Дополнительные сведения см. в статье [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="1fcf8-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="1fcf8-121">Публикация надстройки</span><span class="sxs-lookup"><span data-stu-id="1fcf8-121">Publish the add-in</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="1fcf8-122">См. также</span><span class="sxs-lookup"><span data-stu-id="1fcf8-122">See also</span></span>

- [<span data-ttu-id="1fcf8-123">Основные принципы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1fcf8-123">Core concepts for Office Add-ins</span></span>](../overview/core-concepts-office-add-ins.md)
- [<span data-ttu-id="1fcf8-124">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1fcf8-124">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="1fcf8-125">Проектирование надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1fcf8-125">Design Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="1fcf8-126">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1fcf8-126">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="1fcf8-127">Публикация надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1fcf8-127">Publish Office Add-ins</span></span>](../publish/publish.md)
