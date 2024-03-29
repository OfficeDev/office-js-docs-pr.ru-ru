---
title: Разработка надстроек Office с помощью Visual Studio
description: Как разрабатывать надстройки Office с помощью Visual Studio.
ms.date: 01/26/2022
ms.localizationpriority: high
ms.openlocfilehash: 52740e16363e3e038269e08a9e50e0f08877db66
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743845"
---
# <a name="develop-office-add-ins-with-visual-studio"></a>Разработка надстроек Office с помощью Visual Studio

В этой статье описано, как использовать Visual Studio для разработки надстроек Office. Если надстройка уже создана, можно перейти к разделу [Разработка надстройки с помощью Visual Studio](#develop-the-add-in-using-visual-studio).

> [!NOTE]
> Вместо Visual Studio можно использовать генератор Yeoman для надстроек Office и VS Code для создания надстройки Office. Дополнительные сведения об этом выборе см. в разделе [Создание надстройки Office](../develop/develop-overview.md#create-an-office-add-in).

## <a name="create-the-add-in-project-using-visual-studio"></a>Создание проекта надстройки с помощью Visual Studio

С помощью Visual Studio можно создавать надстройки Office для Excel, Outlook, Word и PowerPoint. Проект надстройки Office создается в рамках решения Visual Studio и использует HTML, CSS и JavaScript. Чтобы создать надстройку Office с помощью Visual Studio, следуйте указаниям из краткого руководства, соответствующего типу надстройки, которую нужно создать.

- [Краткое руководство по началу работы с Excel](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Краткое руководство по началу работы с Outlook](../quickstarts/outlook-quickstart.md?tabs=visualstudio)
- [Краткое руководство по началу работы с Word](../quickstarts/word-quickstart.md?tabs=visualstudio)
- [Краткое руководство по началу работы с PowerPoint](../quickstarts/powerpoint-quickstart.md?tabs=visualstudio)

Visual Studio не поддерживает создание надстроек Office для OneNote или Project. Чтобы создавать надстройки Office для любого из этих приложений, потребуется использовать генератор Yeoman для надстроек Office, как описано в [кратком руководстве по началу работы с OneNote](../quickstarts/onenote-quickstart.md) и в [кратком руководстве по началу работы с Project](../quickstarts/project-quickstart.md).

## <a name="develop-the-add-in-using-visual-studio"></a>Разработка надстройки с помощью Visual Studio

В Visual Studio создается простая надстройка с ограниченными возможностями. Можно настроить надстройку, отредактировав файлы [манифеста](add-in-manifests.md), HTML, JavaScript и CSS в Visual Studio. Общее описание структуры проекта и файлов в проекте надстройки, создаваемом в Visual Studio, см. в справочнике по Visual Studio в составе краткого руководства по началу работы, с помощью которого вы создали надстройку.

> [!TIP]
> Надстройка Office представляет собой веб-приложение, поэтому для изменения надстройки требуются базовые навыки веб-разработки. Если вы впервые работаете с JavaScript, рекомендуем прочесть [учебник Mozilla по JavaScript](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction).

Чтобы настроить надстройку, вам нужно будет усвоить принципы, описанные в разделе [Основные принципы > Разработка](develop-overview.md) этой документации, а также принципы, описанные в соответствующем разделе документации приложения, для которого вы создаете надстройку (например, [Excel](../excel/index.yml)).

## <a name="test-and-debug-the-add-in"></a>Тестирование и отладка надстройки

Методы тестирования, отладки и устранения неполадок надстроек Office зависят от платформы. Дополнительные сведения см. в статьях [Отладка надстроек Office в Visual Studio](debug-office-add-ins-in-visual-studio.md) и [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md).

## <a name="publish-the-add-in"></a>Публикация надстройки

Надстройка Office состоит из веб-приложения и файла манифеста. Веб-приложение определяет пользовательский интерфейс и функции надстройки, а манифест указывает расположение веб-приложения и определяет параметры и возможности надстройки.

В процессе разработки надстройки в Visual Studio эта надстройка запускается на локальном веб-сервере (`localhost`). Если надстройка работает нужным образом и вы готовы опубликовать ее для доступа других пользователей, выполните следующие действия.

1. Разверните веб-приложение на веб-сервере или в службе веб-хостинга (например, Microsoft Azure).
2. Обновите манифест, указав URL-адрес развернутого приложения.
3. Выберите метод [развертывания надстройки Office](../publish/publish.md) и следуйте инструкциям, чтобы опубликовать файл манифеста.

## <a name="see-also"></a>См. также

- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Проектирование надстроек Office](../design/add-in-design.md)
- [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
- [Публикация надстроек Office](../publish/publish.md)
