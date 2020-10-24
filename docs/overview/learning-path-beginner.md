---
title: Руководство для начинающих
description: Рекомендуемый для начинающих путь, включающий использование учебных ресурсов для надстроек Office.
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: a51ffc437c9d1946b886d1e665836dd6d76f52d2
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741073"
---
# <a name="beginners-guide"></a>Руководство для начинающих

Хотите начать создавать собственные кроссплатформенные расширения Office? Следующие шаги покажут вам, что читать в первую очередь, какие инструменты установить и какие учебные пособия рекомендуется выполнить.

> [!NOTE]
> Если у вас есть опыт создания надстроек VSTO для Office, рекомендуем сразу перейти к статье [Руководство для разработчиков надстроек VSTO](learning-path-transition.md), которая дополняет сведения, приведенные в этой статье.

## <a name="step-0-prerequisites"></a>Шаг 0. Необходимые условия

- Надстройки Office - это веб-приложения, встроенные в Office. Итак, сначала вы должны иметь общее представление о веб-приложениях и о том, как они размещаются в сети. Об этом огромное количество информации в Интернете, книгах и онлайн-курсах. Хороший способ начать, если у вас нет предварительных знаний о веб-приложениях, - это поиск "Что такое веб-приложение?" в Bing.
- Основной язык программирования, который вы будете использовать при создании надстроек Office, - это JavaScript или TypeScript. Вы можете думать о TypeScript как о строго типизированной версии JavaScript. Если вы не знакомы ни с одним из этих языков, но у вас есть опыт работы с VBA, VB.Net, C#, вам, вероятно, будет легче освоить TypeScript. Опять же, есть много информации об этих языках в Интернете, книгах и онлайн-курсах.

## <a name="step-1-begin-with-fundamentals"></a>Шаг 1. Начните с основ

Мы знаем, что вам не терпится начать программирование, но есть некоторые вещи о надстройках Office, которые вы должны прочитать, прежде чем открывать свою IDE или редактор кода.

- [Обзор платформы надстроек Office](office-add-ins.md): узнайте, что такое надстройки Office Web и чем они отличаются от более старых способов расширения Office, таких как надстройки VSTO.
- [Разработка надстроек Office](../develop/develop-overview.md). Ознакомьтесь с обзором разработки и жизненного цикла надстроек Office, включая инструменты, создание пользовательского интерфейса надстройки и использование API-интерфейсов JavaScript для взаимодействия с документом Office.

В этих статьях много ссылок, но если вы новичок в надстройках Office, мы рекомендуем вам вернуться сюда после прочтения и перейти к следующему разделу.

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>Шаг 2. Установите инструменты и создайте свою первую надстройку.

Теперь у вас есть общая картина, так что погрузитесь с одним из наших быстрых стартов. В целях изучения платформы мы рекомендуем быстрый запуск Excel. Существует версия, основанная на Visual Studio, и версия, основанная на Node.js и Visual Studio Code.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js и Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>Шаг 3. Код

Вы не можете научиться водить, читая руководство пользователя, поэтому начните программировать с этого [учебника Excel](../tutorials/excel-tutorial.md). Вы будете использовать библиотеку Office JavaScript и немного XML в манифесте надстроек. Нет необходимости запоминать что-либо, потому что на следующих шагах вы получите больше информации об обоих.

## <a name="step-4-understand-the-javascript-library"></a>Шаг 4. Знакомство с библиотекой JavaScript

Во-первых, вы можете получить общее представление о библиотеке JavaScript Office с этим учебным пособием от Microsoft Learn: [Понимание API-интерфейсов Office JavaScript](https://docs.microsoft.com/learn/modules/understand-office-javascript-apis/index).

Затем изучите API-интерфейсы Office JavaScript с помощью нашего [инструмента Script Lab](explore-with-script-lab.md) - песочницы для запуска и изучения API-интерфейсов.

## <a name="step-5-understand-the-manifest"></a>Шаг 5. Знакомство с манифестом

Получите представление о целях манифеста надстройки и ознакомьтесь с его разметкой XML в [манифесте надстроек Office XML](../develop/add-in-manifests.md).

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем с окончанием курса обучения начинающих для надстроек Office! Вот несколько предложений для дальнейшего изучения нашей документации:

- Учебные материалы и краткое руководство для других приложений Office.

  - [Руководство по началу работы с OneNote](../quickstarts/onenote-quickstart.md)
  - [Учебник по Outlook](/outlook/add-ins/addin-tutorial)
  - [Учебник по PowerPoint](../tutorials/powerpoint-tutorial.md)
  - [Руководство по началу работы с Project](../quickstarts/project-quickstart.md)
  - [Учебник по Word](../tutorials/word-tutorial.md)

- Другие важные темы:

  - [Разработка надстроек Office](../develop/develop-overview.md)
  - [Рекомендации по разработке надстроек Office](../concepts/add-in-development-best-practices.md)
  - [Проектирование надстроек Office](../design/add-in-design.md)
  - [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
  - [Развертывание и публикация надстроек Office](../publish/publish.md)
  - [Ресурсы](../resources/resources-links-help.md)
  - [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)