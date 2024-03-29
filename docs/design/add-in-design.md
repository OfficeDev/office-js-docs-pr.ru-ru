---
title: Разработка пользовательского интерфейса надстроек Office
description: Изучите лучшие методики визуального проектирования надстроек Office.
ms.date: 07/08/2021
ms.localizationpriority: high
ms.openlocfilehash: efbb0ee5f0ba75170b8bd4343392c07d9eda8501
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244753"
---
# <a name="design-the-ui-of-office-add-ins"></a>Разработка пользовательского интерфейса надстроек Office

Надстройки Office расширяют возможности Office, предоставляя контекстные функции, с которыми пользователи могут работать в клиентах Office. Надстройки предоставляют доступ ко внешним функциям в Office без необходимости переключаться на другие приложения, что отнимает много времени.

Дизайн интерфейса надстройки должен органично интегрироваться в Office, чтобы пользователи могли эффективно и легко с ним работать. Настройте [команды надстроек](add-in-commands.md) для представления доступа к надстройке и воспользуйтесь нашими рекомендациями при создании пользовательского интерфейса на основе HTML.

## <a name="office-design-principles"></a>Принципы оформления Office

Приложения Office соответствуют общему набору правил взаимодействия. Приложения имеют общий контент и элементы, которые выглядят и работают одинаковым образом. Это сходство основывается на наборе принципов разработки. Принципы помогают команде Office создавать интерфейсы, которые отвечают задачам клиентов. Их понимание и соблюдение поможет вам создавать решения, которые отвечают целям ваших клиентам в Office.

Соблюдайте принципы оформления Office, чтобы ваши надстройки не вызывали у пользователей никаких неудобств.

- **Разрабатывайте специально для Office.** Функциональность, внешний вид и удобство использования надстройки должны гармонично дополнять возможности Office. Надстройки должны соответствовать дизайну Office. Они должны легко интегрироваться в Word на iPad или PowerPoint в Интернете. Хорошая надстройка — это органичное сочетание дополнительных возможностей, платформы и приложения Office. Применяйте тематическое оформление документов и пользовательского интерфейса, где это необходимо. Рекомендуем использовать [Fluent UI для Интернета](https://developer.microsoft.com/fluentui#/get-started/web) в качестве языка дизайна и набора инструментов. В случае Fluent UI для Интернета предусмотрено две версии.

  - **Для пользовательских интерфейсов не на React:** Используйте **Fabric Core**, коллекцию классов CSS с открытым исходным кодом и примесей SASS, обеспечивающих доступ к цветам, анимации, шрифтам, значкам и сеткам. (По историческим причинам он называется "Fabric Core", а не "Fluent Core".) Для начала см. раздел [Fabric Core в надстройках Office](fabric-core.md).
  - **Для пользовательских интерфейсов на React:** используйте **Fluent UI React**, платформу внешнего интерфейса React, предназначенную для создания интерфейсов, которые легко вписываются в широкий спектр продуктов Microsoft. Он обеспечивает надежные, современные, доступные компоненты на основе React, которые легко настраиваются с помощью CSS-in-JS. Для начала, см. раздел [Fluent UI React в надстройках Office](using-office-ui-fabric-react.md).

- **Содержимое важнее, чем хром.** При работе с надстройкой внимание клиентов должно оставаться на странице, слайде или электронной таблице. Надстройка — это вспомогательный интерфейс. Вспомогательный хром не должен мешать работе с содержимым и функциями надстройки. Размещение фирменной символики требует разумного подхода. Мы знаем, что важно сделать надстройку уникальной и узнаваемой, но фирменная символика не должна отвлекать пользователей. Стремитесь к тому, чтобы основное внимание уделялось содержимому и выполнению задач, а не символике.

- **Сделайте работу с надстройкой приятной и разрешите пользователям самим выбирать, что делать.** Людям нравятся функциональные и красивые продукты. Тщательно проработайте свою надстройку. Уделите особое внимание мелочам, учитывайте все варианты взаимодействия и внешнего вида. Разрешите пользователям самим выбирать, что делать. Действия, необходимые для выполнения задачи, должны быть понятными и логичными. Важные решения должны быть понятными. Отмена действий не должна вызывать затруднений. Надстройка — это не конечная точка, а улучшение функциональности Office.

- **Поддержка всех платформ и способов ввода**. Надстройки предназначены для работы на всех платформах, поддерживаемых Office, поэтому интерфейс вашей надстройки должен быть оптимизирован для различных платформ и форм-факторов. Реализуйте поддержку клавиатуры, мыши и сенсорных устройств ввода, а также убедитесь, что пользовательский интерфейс на основе HTML адаптируется к разным форм-факторам. Дополнительные сведения см. в статье [Сенсорный ввод](../concepts/add-in-development-best-practices.md#optimize-for-touch).

## <a name="see-also"></a>См. также

- [Рекомендации по разработке надстроек](../concepts/add-in-development-best-practices.md)
