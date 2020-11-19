---
title: Рекомендации по выбору стиля визуализации данных для надстроек Office
description: Ознакомьтесь с рекомендациями по визуализации данных в надстройке Office.
ms.date: 01/14/2019
localization_priority: Normal
ms.openlocfilehash: f3fa2a6cc5a9d27135ad4290eded838dfaecb7d6
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132076"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="f8704-103">Рекомендации по выбору стиля визуализации данных для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="f8704-103">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="f8704-p101">Качественная визуализация помогает пользователям анализировать данные. Благодаря этому они смогут рассказывать содержательные и убедительные истории. В этой статье представлены рекомендации по эффективной визуализации данных в надстройках для Excel и других приложений Office.</span><span class="sxs-lookup"><span data-stu-id="f8704-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="f8704-p102">Рекомендуем использовать [Office UI Fabric](https://developer.microsoft.com/fabric) при создании хрома для визуализации данных. Office UI Fabric включает стили и компоненты, которые отлично сочетаются с внешним видом Office.</span><span class="sxs-lookup"><span data-stu-id="f8704-p102">We recommend that you use [Office UI Fabric](https://developer.microsoft.com/fabric) to create the chrome for your data visualizations. Office UI Fabric includes styles and components that integrate seamlessly with the Office look and feel.</span></span>
<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a><span data-ttu-id="f8704-109">Элементы визуализации данных</span><span class="sxs-lookup"><span data-stu-id="f8704-109">Data visualization elements</span></span>

<span data-ttu-id="f8704-110">Визуализации данных совместно используют общие платформы и общие визуальные и интерактивные элементы, в том числе заголовки, метки и графики данных, как показано на следующем рисунке.</span><span class="sxs-lookup"><span data-stu-id="f8704-110">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figure.</span></span>

![График с надписями "заголовок", "оси", "условные обозначения" и "область построения"](../images/excel-charts-visualization.png)

### <a name="chart-titles"></a><span data-ttu-id="f8704-112">Заголовки диаграмм</span><span class="sxs-lookup"><span data-stu-id="f8704-112">Chart titles</span></span>

<span data-ttu-id="f8704-113">При создании заголовков диаграмм следуйте таким рекомендациям:</span><span class="sxs-lookup"><span data-stu-id="f8704-113">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="f8704-p103">Сделайте заголовки диаграмм удобочитаемыми. Располагайте их с соблюдением четкой визуальной иерархии относительно остальных элементов диаграммы.</span><span class="sxs-lookup"><span data-stu-id="f8704-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="f8704-p104">Как правило, следует начинать предложения с прописной буквы. Чтобы создать контраст или обозначить иерархию, можно использовать все прописные буквы, но этим не следует злоупотреблять.</span><span class="sxs-lookup"><span data-stu-id="f8704-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="f8704-p105">Используйте [набор шрифтов Office UI Fabric](https://developer.microsoft.com/fabric#/styles/typography), чтобы внешний вид диаграмм сочетался с пользовательским интерфейсом Office, где используется шрифт Segoe. Если же требуется отделить содержимое диаграммы от пользовательского интерфейса, вы можете использовать другой шрифт.</span><span class="sxs-lookup"><span data-stu-id="f8704-p105">Incorporate the [Office UI Fabric type ramp](https://developer.microsoft.com/fabric#/styles/typography) to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="f8704-120">Используйте шрифты sans-serif больших размеров.</span><span class="sxs-lookup"><span data-stu-id="f8704-120">Use sans-serif typefaces with large counters.</span></span>

### <a name="axis-labels"></a><span data-ttu-id="f8704-121">Подписи осей</span><span class="sxs-lookup"><span data-stu-id="f8704-121">Axis labels</span></span>

<span data-ttu-id="f8704-p106">Сделайте подписи осей достаточно темными, чтобы их было легко прочитать. При этом соблюдайте контраст между цветами текста и фона. Убедитесь, что они не настолько темные, чтобы отвлекать внимание от данных.</span><span class="sxs-lookup"><span data-stu-id="f8704-p106">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="f8704-124">Для меток осей лучше всего подходят светло-серые тона.</span><span class="sxs-lookup"><span data-stu-id="f8704-124">Light grays are most effective for axis labels.</span></span> <span data-ttu-id="f8704-125">Если вы используете Fabric, просмотрите [палитру нейтральные цвета](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="f8704-125">If you're using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

### <a name="data-ink"></a><span data-ttu-id="f8704-126">Точки данных</span><span class="sxs-lookup"><span data-stu-id="f8704-126">Data ink</span></span>

<span data-ttu-id="f8704-p108">Пиксели, представляющие фактические данные на диаграмме, называются точками данных. Основное внимание в визуализации должно уделяться им. Не рекомендуется использовать тени, жирные контуры и лишние элементы оформления, которые искажают данные или отвлекают от них внимание. Используйте градиенты, только если значения данных связаны со значениями цветов. Старайтесь не использовать трехмерные диаграммы, если к третьей оси не привязано измеримое целевое значение.</span><span class="sxs-lookup"><span data-stu-id="f8704-p108">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="f8704-132">Цвет</span><span class="sxs-lookup"><span data-stu-id="f8704-132">Color</span></span>

<span data-ttu-id="f8704-p109">Выбирайте цвета, соответствующие темам операционной системы и приложения, а не жестко заданные значения. В то же время убедитесь, что применяемые цвета не искажают данные. Неправильное использование цветов при визуализации данных может привести к искажению данных и неправильному их толкованию.</span><span class="sxs-lookup"><span data-stu-id="f8704-p109">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="f8704-136">Рекомендации по использованию цветов при визуализации данных см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="f8704-136">For best practices for use of color in data visualizations, see the following:</span></span>

- [<span data-ttu-id="f8704-137">Почему цвета радуги — не лучший вариант для визуализации данных</span><span class="sxs-lookup"><span data-stu-id="f8704-137">Why rainbow colors aren't the best option for data visualizations</span></span>](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="f8704-138">Color Brewer 2.0: советы по выбору цветов для картографии</span><span class="sxs-lookup"><span data-stu-id="f8704-138">Color Brewer 2.0: Color Advice for Cartography</span></span>](https://colorbrewer2.org/)
- [<span data-ttu-id="f8704-139">Как выбрать оттенок</span><span class="sxs-lookup"><span data-stu-id="f8704-139">I Want Hue</span></span>](https://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="f8704-140">Линии сетки</span><span class="sxs-lookup"><span data-stu-id="f8704-140">Gridlines</span></span>

<span data-ttu-id="f8704-p110">Как правило, линии сетки необходимы для точного чтения диаграммы, но их можно представить как вспомогательный визуальный элемент, который выделяет точки данных, а не отвлекает от них. Сделайте статические линии сетки тонкими и светлыми, если они не создаются специально для усиления контраста. Вы также можете создать динамические линии сетки, своевременно появляющиеся в зависимости от контекста, в котором пользователь работает с диаграммой.</span><span class="sxs-lookup"><span data-stu-id="f8704-p110">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="f8704-144">Для линий сетки лучше всего подходят светло-серые тона.</span><span class="sxs-lookup"><span data-stu-id="f8704-144">Light grays are most effective for gridlines.</span></span> <span data-ttu-id="f8704-145">Если вы используете Fabric, просмотрите [палитру нейтральные цвета](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="f8704-145">If you're using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

<span data-ttu-id="f8704-146">На приведенном ниже рисунке показана визуализация данных с линиями сетки.</span><span class="sxs-lookup"><span data-stu-id="f8704-146">The following image shows a data visualization with gridlines.</span></span>

![Визуализация данных графика с линиями сетки](../images/data-visualization.png)

### <a name="legends"></a><span data-ttu-id="f8704-148">Условные обозначения</span><span class="sxs-lookup"><span data-stu-id="f8704-148">Legends</span></span>

<span data-ttu-id="f8704-149">Условные обозначения необходимы для следующего:</span><span class="sxs-lookup"><span data-stu-id="f8704-149">Add legends if necessary to:</span></span>

- <span data-ttu-id="f8704-150">различения рядов данных;</span><span class="sxs-lookup"><span data-stu-id="f8704-150">Distinguish between series</span></span>
- <span data-ttu-id="f8704-151">представления изменений масштаба и значений.</span><span class="sxs-lookup"><span data-stu-id="f8704-151">Present scale or value changes</span></span>

<span data-ttu-id="f8704-p112">Убедитесь, что условные обозначения выделяют точки данных, а не отвлекают от них. Располагайте условные обозначения следующим образом:</span><span class="sxs-lookup"><span data-stu-id="f8704-p112">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="f8704-154">С выравниванием по левому краю над областью представления данных по умолчанию, если все обозначения помещаются над диаграммой.</span><span class="sxs-lookup"><span data-stu-id="f8704-154">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="f8704-155">Справа вверху в области представления данных, если все обозначения не помещаются над диаграммой. При необходимости можно разрешить прокрутку списка.</span><span class="sxs-lookup"><span data-stu-id="f8704-155">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="f8704-p113">Для наглядности придайте маркерам условных обозначений форму, соответствующую типу диаграммы. Например, круглые маркеры подходят для точечных и пузырьковых диаграмм. Для графиков подходят маркеры в виде сегментов линий.</span><span class="sxs-lookup"><span data-stu-id="f8704-p113">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="f8704-159">Подписи и подсказки данных</span><span class="sxs-lookup"><span data-stu-id="f8704-159">Data labels and tooltips</span></span>

<span data-ttu-id="f8704-p114">Убедитесь, что в подписях и подсказках данных используются достаточно большие отступы и подходящие типы. Используйте алгоритмы, чтобы свести к минимуму наложения. Например, всплывающая подсказка может по умолчанию появляться справа от данных, если соответствующая точка не находится слишком близко к правому краю.</span><span class="sxs-lookup"><span data-stu-id="f8704-p114">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="f8704-163">Принципы оформления</span><span class="sxs-lookup"><span data-stu-id="f8704-163">Design principles</span></span>

<span data-ttu-id="f8704-164">Команда разработчиков Office составила приведенный ниже список принципов оформления, которым мы следуем при визуализации данных для набора продуктов Office.</span><span class="sxs-lookup"><span data-stu-id="f8704-164">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="f8704-165">Принципы визуального оформления</span><span class="sxs-lookup"><span data-stu-id="f8704-165">Visual design principles</span></span>

- <span data-ttu-id="f8704-p115">Визуализация должна точно и качественно передавать данные, чтобы их было легче понять. Выделяйте данные с помощью вспомогательных элементов только в той степени, которой требует контекст. Избегайте лишних украшений (теней, контуров и т. д.), ненужных элементов и искажения данных.</span><span class="sxs-lookup"><span data-stu-id="f8704-p115">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="f8704-p116">Визуализация должна вызывать интерес за счет наглядных зрительных образов. Используйте традиционные шаблоны взаимодействия, элементы управления и понятные реакции системы.</span><span class="sxs-lookup"><span data-stu-id="f8704-p116">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="f8704-p117">Применяйте проверенные временем принципы оформления. Следуйте традиционным принципам типографии и визуальной передачи, чтобы улучшить оформление, повысить удобочитаемость и точно передать смысл.</span><span class="sxs-lookup"><span data-stu-id="f8704-p117">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="f8704-173">Принципы взаимодействия</span><span class="sxs-lookup"><span data-stu-id="f8704-173">Interaction design principles</span></span>

- <span data-ttu-id="f8704-174">Диаграмма должна вызывать интерес.</span><span class="sxs-lookup"><span data-stu-id="f8704-174">Design to allow for exploration.</span></span>
- <span data-ttu-id="f8704-175">Обеспечьте непосредственное взаимодействие с объектами, позволяющее взглянуть на данные с новой стороны (например, сортировку путем перетаскивания).</span><span class="sxs-lookup"><span data-stu-id="f8704-175">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="f8704-176">Используйте простые, непосредственные и привычные модели взаимодействия.</span><span class="sxs-lookup"><span data-stu-id="f8704-176">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="f8704-177">Дополнительные сведения о создании понятных интерактивных представлений данных см. в статье [Принципы и распространенные ошибки оформления интерфейса](https://uitraps.com/).</span><span class="sxs-lookup"><span data-stu-id="f8704-177">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](https://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="f8704-178">Принципы динамического оформления</span><span class="sxs-lookup"><span data-stu-id="f8704-178">Motion design principles</span></span>

<span data-ttu-id="f8704-p118">Движение — результат воздействия. Визуальные элементы должны двигаться в одном направлении и с одинаковой скоростью. Это относится к следующему:</span><span class="sxs-lookup"><span data-stu-id="f8704-p118">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="f8704-182">созданию диаграмм;</span><span class="sxs-lookup"><span data-stu-id="f8704-182">Chart creation</span></span>
- <span data-ttu-id="f8704-183">изменению типа диаграммы;</span><span class="sxs-lookup"><span data-stu-id="f8704-183">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="f8704-184">фильтрам;</span><span class="sxs-lookup"><span data-stu-id="f8704-184">Filtering</span></span>
- <span data-ttu-id="f8704-185">сортировке;</span><span class="sxs-lookup"><span data-stu-id="f8704-185">Sorting</span></span>
- <span data-ttu-id="f8704-186">сложению и вычитанию данных;</span><span class="sxs-lookup"><span data-stu-id="f8704-186">Adding or subtracting data</span></span>
- <span data-ttu-id="f8704-187">объединению и сегментации данных;</span><span class="sxs-lookup"><span data-stu-id="f8704-187">Brushing or slicing data</span></span>
- <span data-ttu-id="f8704-188">изменению размера диаграммы;</span><span class="sxs-lookup"><span data-stu-id="f8704-188">Resizing a chart</span></span>

<span data-ttu-id="f8704-p119">созданию ощущения непринужденности. При создании анимации следуйте таким рекомендациям:</span><span class="sxs-lookup"><span data-stu-id="f8704-p119">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="f8704-191">Проектируйте элементы по одному.</span><span class="sxs-lookup"><span data-stu-id="f8704-191">Stage one thing at a time.</span></span>
- <span data-ttu-id="f8704-192">Изменяйте оси, прежде чем менять точки данных.</span><span class="sxs-lookup"><span data-stu-id="f8704-192">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="f8704-193">Если объекты двигаются в одном направлении и с одинаковой скоростью, обрабатывайте их как группу.</span><span class="sxs-lookup"><span data-stu-id="f8704-193">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="f8704-p120">Собирайте элементы в группы не более чем из 4–5 объектов. Пользователям сложно отслеживать более 4–5 независимых объектов.</span><span class="sxs-lookup"><span data-stu-id="f8704-p120">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="f8704-196">Движение добавляет осмысленность.</span><span class="sxs-lookup"><span data-stu-id="f8704-196">Motion adds meaning.</span></span>

- <span data-ttu-id="f8704-197">Анимация помогает пользователям ориентироваться в изменениях данных, создает контекст и заменяет комментарии.</span><span class="sxs-lookup"><span data-stu-id="f8704-197">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="f8704-198">Движение должно происходить в понятном координатном пространстве визуализации.</span><span class="sxs-lookup"><span data-stu-id="f8704-198">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="f8704-199">Анимация должна соответствовать визуальному оформлению.</span><span class="sxs-lookup"><span data-stu-id="f8704-199">Tailor the animation to the visual.</span></span>
- <span data-ttu-id="f8704-200">Не используйте анимацию без необходимости.</span><span class="sxs-lookup"><span data-stu-id="f8704-200">Avoid gratuitous animations.</span></span>

<span data-ttu-id="f8704-201">Движение следует за данными.</span><span class="sxs-lookup"><span data-stu-id="f8704-201">Motion follows data.</span></span>

- <span data-ttu-id="f8704-p121">Сохраняйте сопоставления данных. Если область привязана к показателю, сохраняйте ее при переходе.</span><span class="sxs-lookup"><span data-stu-id="f8704-p121">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="f8704-p122">Все анимации должны быть выдержаны в одном стиле. По возможности согласуйте анимацию визуализации данных с оформлением Office. Используйте аналогичные анимации для похожих диаграмм.</span><span class="sxs-lookup"><span data-stu-id="f8704-p122">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="f8704-207">Специальные возможности для визуализации данных</span><span class="sxs-lookup"><span data-stu-id="f8704-207">Accessibility in data visualizations</span></span>

- <span data-ttu-id="f8704-p123">Цвет не должен быть единственным способом передачи информации. В противном случае люди, страдающие дальтонизмом, не смогут толковать результаты. По мере возможности используйте для передачи информации не только цвет, но и форму, размер и текстуры.</span><span class="sxs-lookup"><span data-stu-id="f8704-p123">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="f8704-211">Обеспечьте возможность управлять с помощью клавиатуры всеми интерактивными элементами, такими как кнопки и списки.</span><span class="sxs-lookup"><span data-stu-id="f8704-211">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="f8704-212">Отправляйте события специальных возможностей средствам чтения с экрана для объявления об изменениях фокуса, всплывающих подсказках и т. д.</span><span class="sxs-lookup"><span data-stu-id="f8704-212">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="f8704-213">См. также</span><span class="sxs-lookup"><span data-stu-id="f8704-213">See also</span></span>

- [<span data-ttu-id="f8704-214">Пять лучших библиотек для визуализации данных</span><span class="sxs-lookup"><span data-stu-id="f8704-214">The Five Best Libraries for Building Data Visualizations</span></span>](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="f8704-215">Визуальное представление количественных данных</span><span class="sxs-lookup"><span data-stu-id="f8704-215">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
