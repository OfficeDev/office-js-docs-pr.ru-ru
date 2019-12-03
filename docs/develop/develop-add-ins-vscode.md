---
title: Разработка надстроек Office с помощью Visual Studio Code
description: Как разрабатывать надстройки Office с помощью Visual Studio Code
ms.date: 12/02/2019
localization_priority: Priority
ms.openlocfilehash: a18d8a74ff269b32e83c836b06629850873e507b
ms.sourcegitcommit: 44f1a4a3e1ae3c33d7d5fabcee14b84af94e03da
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/03/2019
ms.locfileid: "39670500"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a><span data-ttu-id="cca7a-103">Разработка надстроек Office с помощью Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="cca7a-103">Develop Office Add-ins with Visual Studio Code</span></span>

<span data-ttu-id="cca7a-104">В этой статье описано, как разработать надстройку Office с помощью [Visual Studio Code (VS Code)](https://code.visualstudio.com).</span><span class="sxs-lookup"><span data-stu-id="cca7a-104">This article describes how to use [Visual Studio Code (VS Code)](https://code.visualstudio.com) to develop an Office Add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="cca7a-105">Сведения о создании надстройки Office с помощью Visual Studio см. в статье [Создание и отладка надстроек Office в Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="cca7a-105">For information about using Visual Studio to create an Office Add-in, see [Create and debug Office Add-ins in Visual Studio](create-and-debug-office-add-ins-in-visual-studio.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="cca7a-106">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="cca7a-106">Prerequisites</span></span>

- [<span data-ttu-id="cca7a-107">Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="cca7a-107">Visual Studio Code</span></span>](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a><span data-ttu-id="cca7a-108">Создание проекта надстройки с помощью генератора Yeoman</span><span class="sxs-lookup"><span data-stu-id="cca7a-108">Create the add-in project using the Yeoman generator</span></span>

<span data-ttu-id="cca7a-109">Если вы используете VS Code в качестве интегрированной среды разработки (IDE), следует создать проект надстройки Office с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Генератор Yeoman создает проект Node.js, которым можно управлять с помощью VS Code или любого другого редактора.</span><span class="sxs-lookup"><span data-stu-id="cca7a-109">If you're using VS Code as your integrated development environment (IDE), you should create the Office Add-in project with the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). The Yeoman generator creates a Node.js project that can be managed with VS Code or any other editor.</span></span> 

<span data-ttu-id="cca7a-110">Чтобы создать надстройку Office с помощью генератора Yeoman, следуйте указаниям из [5-минутного краткого руководства](../index.md), соответствующего типу надстройки, которую нужно создать.</span><span class="sxs-lookup"><span data-stu-id="cca7a-110">To create an Office Add-in with the Yeoman generator, follow instructions in the [5-minute quick start](../index.md) that corresponds to the type of add-in you'd like to create.</span></span>

## <a name="develop-the-add-in-using-vs-code"></a><span data-ttu-id="cca7a-111">Разработка надстройки с помощью VS Code</span><span class="sxs-lookup"><span data-stu-id="cca7a-111">Develop the add-in using VS Code</span></span>

<span data-ttu-id="cca7a-112">Когда генератор Yeoman закончит создание проекта надстройки, откройте корневую папку проекта с помощью VS Code.</span><span class="sxs-lookup"><span data-stu-id="cca7a-112">When the Yeoman generator finishes creating the add-in project, open the root folder of the project with VS Code.</span></span> 

> [!TIP]
> <span data-ttu-id="cca7a-113">В Windows вы можете перейти в корневой каталог проекта с помощью командной строки и ввести `code .`, чтобы открыть эту папку в VS Code.</span><span class="sxs-lookup"><span data-stu-id="cca7a-113">On Windows, you can navigate to the root directory of the project via the command line and then enter `code .` to open that folder in VS Code.</span></span> <span data-ttu-id="cca7a-114">На компьютере Mac потребуется [добавить в путь команду `code`](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) перед использованием этой команды для открытия папки проекта в VS Code.</span><span class="sxs-lookup"><span data-stu-id="cca7a-114">On Mac, you'll need to [add the `code` command to the path](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) before you can use that command to open the project folder in VS Code.</span></span>

<span data-ttu-id="cca7a-115">Генератор Yeoman создает простую надстройку с ограниченными возможностями.</span><span class="sxs-lookup"><span data-stu-id="cca7a-115">The Yeoman generator creates a basic add-in with limited functionality.</span></span> <span data-ttu-id="cca7a-116">Вы можете настроить надстройку, изменив файлы [манифеста](add-in-manifests.md), HTML, JavaScript, TypeScript или CSS в VS Code.</span><span class="sxs-lookup"><span data-stu-id="cca7a-116">You can customize the add-in by editing the [manifest](add-in-manifests.md), HTML, JavaScript or TypeScript, and CSS files in VS Code.</span></span> <span data-ttu-id="cca7a-117">Общее описание структуры проекта и файлов в проекте надстройки, созданном генератором Yeoman, см. в рекомендациях по генератору Yeoman в [5-минутном кратком руководстве](../index.md), соответствующем типу созданной надстройки.</span><span class="sxs-lookup"><span data-stu-id="cca7a-117">For a high-level description of the project structure and files in the add-in project that the Yeoman generator creates, see the Yeoman generator guidance within the [5-minute quick start](../index.md) that corresponds to the type of add-in you've created.</span></span>

## <a name="test-and-debug-the-add-in"></a><span data-ttu-id="cca7a-118">Тестирование и отладка надстройки</span><span class="sxs-lookup"><span data-stu-id="cca7a-118">To run and debug the add-in</span></span>

<span data-ttu-id="cca7a-119">Методы тестирования, отладки и устранения неполадок надстроек Office зависят от платформы.</span><span class="sxs-lookup"><span data-stu-id="cca7a-119">Methods for testing, debugging, and troubleshooting Office Add-ins vary by platform.</span></span> <span data-ttu-id="cca7a-120">Дополнительные сведения см. в статье [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="cca7a-120">For more information, see [Test and debug Office Add-ins](../testing/test-debug-office-add-ins.md).</span></span>

## <a name="publish-the-add-in"></a><span data-ttu-id="cca7a-121">Публикация надстройки</span><span class="sxs-lookup"><span data-stu-id="cca7a-121">Publish the add-in.</span></span>

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a><span data-ttu-id="cca7a-122">См. также</span><span class="sxs-lookup"><span data-stu-id="cca7a-122">See also</span></span>

- [<span data-ttu-id="cca7a-123">5-минутные краткие руководства</span><span class="sxs-lookup"><span data-stu-id="cca7a-123">5-Minute Quick Starts</span></span>](../index.md)
- <span data-ttu-id="cca7a-124">[Изучение API JavaScript для Office с помощью Script Lab](../overview/explore-with-script-lab.md)</span><span class="sxs-lookup"><span data-stu-id="cca7a-124">To learn more, see [Explore Office JavaScript API using Script Lab](../overview/explore-with-script-lab.md).</span></span>
- [<span data-ttu-id="cca7a-125">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="cca7a-125">Test and debug Office Add-ins</span></span>](../testing/test-debug-office-add-ins.md)
- [<span data-ttu-id="cca7a-126">Развертывание и публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="cca7a-126">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)