---
title: Руководство для разработчиков надстроек VSTO
description: Рекомендуемый путь для опытных разработчиков надстроек VSTO по изучению ресурсов веб-надстроек Office.
ms.date: 10/14/2020
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: bc27177c67028e57030c9baed6b416d0c57c77d1
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810129"
---
# <a name="vsto-add-in-developers-guide"></a>Руководство для разработчиков надстроек VSTO

Итак, вы создали несколько надстроек VSTO для приложений Office, работающих в Windows, и теперь вы изучаете новый способ расширения Office, который будет работать в Windows, Mac и версии веб-браузера набора Office: Веб-надстройки Office.

Your understanding of the object models for the Excel, Word, and the other Office applications will be a huge help because the object models in Office Web Add-ins follow similar patterns. But there are going to be some challenges:

- Вы будете работать с другим языком (JavaScript или TypeScript) вместо C# или Visual Basic .NET. (Также существует описанный ниже способ повторного использования в веб-надстройке некоторых фрагментов вашего существующего кода.)
- Развертывание веб-надстроек Office отличается от развертывания надстроек VSTO.
- Веб-надстройки Office — это веб-приложения, работающие в упрощенном окне браузера, внедренном в приложение Office, поэтому вы должны понимать основы веб-приложений, а также способ их размещения на веб-серверах или в облачных учетных записях. 

По этим причинам значительная часть этой статьи повторяет схему обучения начинающих работе с расширениями Office: [Руководство для начинающих](learning-path-beginner.md). Мы добавили дополнительные учебные материалы, которые помогут разработчикам надстроек VSTO применить свой опыт и повторно использовать свой готовый код.

## <a name="step-0-prerequisites"></a>Шаг 0. Необходимые условия

- Веб-надстройки Office (другое название — надстройки Office) по сути являются веб-приложениями, внедренными в Office. Итак, сначала вы должны иметь общее представление о веб-приложениях и о том, как они размещаются в сети. Об этом доступно огромное количество информации в Интернете, книгах и онлайн-курсах. Хороший способ начать, если у вас нет предварительных знаний о веб-приложениях, - это поиск "Что такое веб-приложение?" в Bing.
- Основным языком программирования, используемым при создании надстроек Office, является JavaScript или TypeScript. Вы можете думать о TypeScript как о строго типизированной версии JavaScript. Если вы не знакомы ни с одним из этих языков, но у вас есть опыт работы с VBA, VB.Net, C#, вам, вероятно, будет легче освоить TypeScript. Опять же, есть много информации об этих языках в Интернете, книгах и онлайн-курсах.

## <a name="step-1-begin-with-fundamentals"></a>Шаг 1. Начните с основ

Мы знаем, что вам не терпится начать программирование, но есть некоторые вещи о надстройках Office, которые вы должны прочитать, прежде чем открывать свою IDE или редактор кода.

- [Обзор платформы надстроек Office](office-add-ins.md): узнайте, что такое надстройки Office Web и чем они отличаются от более старых способов расширения Office, таких как надстройки VSTO.
- [Разработка надстроек Office](../develop/develop-overview.md). Ознакомьтесь с обзором разработки и жизненного цикла надстроек Office, включая инструменты, создание пользовательского интерфейса надстройки и использование API-интерфейсов JavaScript для взаимодействия с документом Office.

В этих статьях есть много ссылок, но если вы переходите к веб-надстройкам Office, рекомендуем вам вернуться сюда после их прочтения и продолжить со следующего раздела.

## <a name="step-2-install-tools-and-create-your-first-add-in"></a>Шаг 2. Установите инструменты и создайте свою первую надстройку.

Теперь у вас есть общая картина, так что погрузитесь с одним из наших быстрых стартов. В целях изучения платформы мы рекомендуем быстрый запуск Excel. Существует версия, основанная на Visual Studio, и другая версия, основанная на Node.js и Visual Studio Code. Если вы переходите с надстроек VSTO, скорее всего, вам будет удобнее работать с версией Visual Studio.

- [Visual Studio](../quickstarts/excel-quickstart-jquery.md?tabs=visualstudio)
- [Node.js и Visual Studio Code](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)

## <a name="step-3-code"></a>Шаг 3. Код

Вы не можете научиться водить, читая руководство пользователя, поэтому начните программировать с этого [учебника Excel](../tutorials/excel-tutorial.md). Вы будете использовать библиотеку Office JavaScript и немного XML в манифесте надстроек. Нет необходимости запоминать что-либо, потому что на следующем шаге будет представлен их дополнительный обзор.

## <a name="step-4-understand-the-javascript-library"></a>Шаг 4. Знакомство с библиотекой JavaScript

Общие сведения о библиотеке JavaScript для Office см. в этом руководстве из учебного курса Microsoft Learn: [Общие сведения об API JavaScript для Office](/training/modules/intro-office-add-ins/3-apis).

Затем изучите API JavaScript для Office с помощью [инструмента Script Lab](explore-with-script-lab.md) — песочницы для запуска и изучения API-интерфейсов.

### <a name="special-resource-for-vsto-add-in-developers"></a>Специальный ресурс для разработчиков надстроек VSTO

Это отличный момент для знакомства с примером надстройки [JavaScript SalesTracker для Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker). Она была создана для выделения похожих элементов и различий между надстройками VSTO и веб-надстройками Office, а в файле сведений примера указываются важные моменты сравнения.

## <a name="step-5-understand-the-manifest"></a>Шаг 5. Знакомство с манифестом

Ознакомьтесь с целями манифеста веб-надстройки и его разметкой XML в [XML-манифесте надстроек Office](../develop/add-in-manifests.md).

## <a name="step-6-for-vsto-developers-only-reuse-your-vsto-code"></a>Шаг 6 (только для разработчиков VSTO). Повторное использование кода VSTO

Вы можете повторно использовать некоторые фрагменты кода надстроек VSTO в веб-надстройках Office, перенося их в серверную часть своего веб-приложения и предоставляя к ним доступ для JavaScript или TypeScript в виде веб-API. В качестве инструкций см. документ [Руководство. Обмен кодом между надстройкой VSTO и надстройкой Office с использованием общей библиотеки кода](../tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial.md).

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем с окончанием схемы обучения разработчиков надстроек VSTO для веб-надстроек Office! Вот несколько предложений для дальнейшего изучения нашей документации:

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
