---
title: Разработка надстроек Office с помощью Visual Studio Code
description: Как разрабатывать надстройки Office с помощью Visual Studio Code
ms.date: 01/16/2020
localization_priority: Priority
ms.openlocfilehash: 0aef01c5b892a0cf08254ca8ffd9dd751b993139
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608319"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a>Разработка надстроек Office с помощью Visual Studio Code

В этой статье описано, как разработать надстройку Office с помощью [Visual Studio Code (VS Code)](https://code.visualstudio.com).

> [!NOTE]
> Сведения об использовании Visual Studio для создания надстроек Office см. в статье [Разработка надстроек Office в Visual Studio](develop-add-ins-visual-studio.md).

## <a name="prerequisites"></a>Необходимые компоненты

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a>Создание проекта надстройки с помощью генератора Yeoman

Если вы используете VS Code в качестве интегрированной среды разработки (IDE), следует создать проект надстройки Office с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Генератор Yeoman создает проект Node.js, которым можно управлять с помощью VS Code или любого другого редактора. 

Чтобы создать надстройку Office с помощью генератора Yeoman, следуйте указаниям из [5-минутного краткого руководства](../index.md), соответствующего типу надстройки, которую нужно создать.

## <a name="develop-the-add-in-using-vs-code"></a>Разработка надстройки с помощью VS Code

Когда генератор Yeoman закончит создание проекта надстройки, откройте корневую папку проекта с помощью VS Code. 

> [!TIP]
> В Windows вы можете перейти в корневой каталог проекта с помощью командной строки и ввести `code .`, чтобы открыть эту папку в VS Code. На компьютере Mac потребуется [добавить в путь команду `code`](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) перед использованием этой команды для открытия папки проекта в VS Code.

Генератор Yeoman создает простую надстройку с ограниченными возможностями. Вы можете настроить надстройку, изменив файлы [манифеста](add-in-manifests.md), HTML, JavaScript, TypeScript или CSS в VS Code. Общее описание структуры проекта и файлов в проекте надстройки, созданном генератором Yeoman, см. в рекомендациях по генератору Yeoman в [5-минутном кратком руководстве](../index.md), соответствующем типу созданной надстройки.

## <a name="test-and-debug-the-add-in"></a>Тестирование и отладка надстройки

Методы тестирования, отладки и устранения неполадок надстроек Office зависят от платформы. Дополнительные сведения см. в статье [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md).

## <a name="publish-the-add-in"></a>Публикация надстройки

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>См. также

- [Создание надстроек Office](../overview/office-add-ins-fundamentals.md)
- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Проектирование надстроек Office](../design/add-in-design.md)
- [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
- [Публикация надстроек Office](../publish/publish.md)