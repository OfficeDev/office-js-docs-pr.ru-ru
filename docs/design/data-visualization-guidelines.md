---
title: Рекомендации по выбору стиля визуализации данных для надстроек Office
description: ''
ms.date: 01/14/2019
localization_priority: Normal
ms.openlocfilehash: ef82432dacb3f63e85fd305bc682325af3312aca
ms.sourcegitcommit: 212c810f3480a750df779777c570159a7f76054a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/17/2020
ms.locfileid: "41217267"
---
# <a name="data-visualization-style-guidelines-for-office-add-ins"></a><span data-ttu-id="db5fa-102">Рекомендации по выбору стиля визуализации данных для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="db5fa-102">Data visualization style guidelines for Office Add-ins</span></span>

<span data-ttu-id="db5fa-p101">Качественная визуализация помогает пользователям анализировать данные. Благодаря этому они смогут рассказывать содержательные и убедительные истории. В этой статье представлены рекомендации по эффективной визуализации данных в надстройках для Excel и других приложений Office.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p101">Good data visualizations help users find insights in their data. They can use those insights to tell stories that inform and persuade. This article provides guidelines to help you design effective data visualizations in your add-ins for Excel and other Office apps.</span></span>

<span data-ttu-id="db5fa-p102">Рекомендуем использовать [Office UI Fabric](https://developer.microsoft.com/fabric) при создании хрома для визуализации данных. Office UI Fabric включает стили и компоненты, которые отлично сочетаются с внешним видом Office.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p102">We recommend that you use [Office UI Fabric](https://developer.microsoft.com/fabric) to create the chrome for your data visualizations. Office UI Fabric includes styles and components that integrate seamlessly with the Office look and feel.</span></span> 

<!--The following figure shows a data visualization in an add-in that uses Fabric.

![Image of a data visualization with Fabric elements applied**](../images/fabric-data-visualization.png) 

-->

## <a name="data-visualization-elements"></a><span data-ttu-id="db5fa-108">Элементы визуализации данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-108">Data visualization elements</span></span>

<span data-ttu-id="db5fa-109">Визуализации данных совместно используют общие платформы и общие визуальные и интерактивные элементы, в том числе заголовки, метки и графики данных, как показано на следующем рисунке.</span><span class="sxs-lookup"><span data-stu-id="db5fa-109">Data visualizations share a general framework and common visual and interactive elements, including titles, labels, and data plots, as shown in the following figure.</span></span>

![Изображение графика с заголовком, осями, условными обозначениями и областью построения с подписью](../images/excel-charts-visualization.png)

### <a name="chart-titles"></a><span data-ttu-id="db5fa-111">Заголовки диаграмм</span><span class="sxs-lookup"><span data-stu-id="db5fa-111">Chart titles</span></span>

<span data-ttu-id="db5fa-112">При создании заголовков диаграмм следуйте таким рекомендациям:</span><span class="sxs-lookup"><span data-stu-id="db5fa-112">Follow these guidelines for chart titles:</span></span>

- <span data-ttu-id="db5fa-p103">Сделайте заголовки диаграмм удобочитаемыми. Располагайте их с соблюдением четкой визуальной иерархии относительно остальных элементов диаграммы.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p103">Make your chart titles easily readable. Position them to create a clear visual hierarchy in relation to the rest of the chart.</span></span>
- <span data-ttu-id="db5fa-p104">Как правило, следует начинать предложения с прописной буквы. Чтобы создать контраст или обозначить иерархию, можно использовать все прописные буквы, но этим не следует злоупотреблять.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p104">In general, use sentence capitalization (capitalize the first word). To create contrast or to reinforce hierarchies, you can use all caps, but all caps should be used sparingly.</span></span>
- <span data-ttu-id="db5fa-p105">Используйте [набор шрифтов Office UI Fabric](https://developer.microsoft.com/fabric#/styles/typography), чтобы внешний вид диаграмм сочетался с пользовательским интерфейсом Office, где используется шрифт Segoe. Если же требуется отделить содержимое диаграммы от пользовательского интерфейса, вы можете использовать другой шрифт.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p105">Incorporate the [Office UI Fabric type ramp](https://developer.microsoft.com/fabric#/styles/typography) to make your charts consistent with the Office UI, which uses Segoe. You can also use a different typeface to differentiate chart content from the UI.</span></span>
- <span data-ttu-id="db5fa-119">Используйте шрифты sans-serif больших размеров.</span><span class="sxs-lookup"><span data-stu-id="db5fa-119">Use sans-serif typefaces with large counters.</span></span>

### <a name="axis-labels"></a><span data-ttu-id="db5fa-120">Подписи осей</span><span class="sxs-lookup"><span data-stu-id="db5fa-120">Axis labels</span></span>

<span data-ttu-id="db5fa-p106">Сделайте подписи осей достаточно темными, чтобы их было легко прочитать. При этом соблюдайте контраст между цветами текста и фона. Убедитесь, что они не настолько темные, чтобы отвлекать внимание от данных.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p106">Make your axis labels dark enough to read clearly, with adequate contrast ratios between the text and background colors. Make sure that they are not so dark that they compete with data ink.</span></span>

<span data-ttu-id="db5fa-p107">Для меток осей лучше всего подходят светло-серые тона. Если вы используете Fabric, см. [нейтральную цветовую палитру](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="db5fa-p107">Light grays are most effective for axis labels. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

### <a name="data-ink"></a><span data-ttu-id="db5fa-125">Точки данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-125">Data ink</span></span>

<span data-ttu-id="db5fa-p108">Пиксели, представляющие фактические данные на диаграмме, называются точками данных. Основное внимание в визуализации должно уделяться им. Не рекомендуется использовать тени, жирные контуры и лишние элементы оформления, которые искажают данные или отвлекают от них внимание. Используйте градиенты, только если значения данных связаны со значениями цветов. Старайтесь не использовать трехмерные диаграммы, если к третьей оси не привязано измеримое целевое значение.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p108">The pixels that represent the actual data in a chart are referred to as data ink. This should be the central focus of the visualization. Avoid the use of drop shadows, heavy outlines, or unnecessary design elements that distort or compete with the data. Use gradients only when data values are tied to color values. Avoid three-dimensional charts unless a measurable, objective value is bound to a third dimension.</span></span>

### <a name="color"></a><span data-ttu-id="db5fa-131">Цвет</span><span class="sxs-lookup"><span data-stu-id="db5fa-131">Color</span></span>

<span data-ttu-id="db5fa-p109">Выбирайте цвета, соответствующие темам операционной системы и приложения, а не жестко заданные значения. В то же время убедитесь, что применяемые цвета не искажают данные. Неправильное использование цветов при визуализации данных может привести к искажению данных и неправильному их толкованию.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p109">Choose colors that follow operating system or application themes rather than hardcoded colors. At the same time, make sure that the colors you apply do not distort the data. Misuse of color in data visualizations can result in data distortion and incorrect reading of information.</span></span>

<span data-ttu-id="db5fa-135">Рекомендации по использованию цветов при визуализации данных см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="db5fa-135">For best practices for use of color in data visualizations, see the following:</span></span>

- [<span data-ttu-id="db5fa-136">Почему цвета радуги — не лучший вариант для визуализации данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-136">Why rainbow colors aren't the best option for data visualizations</span></span>](https://www.poynter.org/2013/why-rainbow-colors-arent-always-the-best-options-for-data-visualizations/224413/)
- [<span data-ttu-id="db5fa-137">Color Brewer 2.0: советы по выбору цветов для картографии</span><span class="sxs-lookup"><span data-stu-id="db5fa-137">Color Brewer 2.0: Color Advice for Cartography</span></span>](http://colorbrewer2.org/)
- [<span data-ttu-id="db5fa-138">Как выбрать оттенок</span><span class="sxs-lookup"><span data-stu-id="db5fa-138">I Want Hue</span></span>](https://tools.medialab.sciences-po.fr/iwanthue/)

### <a name="gridlines"></a><span data-ttu-id="db5fa-139">Линии сетки</span><span class="sxs-lookup"><span data-stu-id="db5fa-139">Gridlines</span></span>

<span data-ttu-id="db5fa-p110">Как правило, линии сетки необходимы для точного чтения диаграммы, но их можно представить как вспомогательный визуальный элемент, который выделяет точки данных, а не отвлекает от них. Сделайте статические линии сетки тонкими и светлыми, если они не создаются специально для усиления контраста. Вы также можете создать динамические линии сетки, своевременно появляющиеся в зависимости от контекста, в котором пользователь работает с диаграммой.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p110">Gridlines are often necessary for accurately reading a chart, but should be presented as a secondary visual element, enhancing the data ink, not competing with it. Make static gridlines thin and light, unless they are designed specifically for high contrast. You can also use interaction to create dynamic, just-in-time gridlines that appear in context when a user interacts with a chart.</span></span>

<span data-ttu-id="db5fa-p111">Для линий сетки лучше всего подходят светло-серые тона. Если вы используете Fabric, см. [нейтральную цветовую палитру](https://developer.microsoft.com/fabric#/styles/colors).</span><span class="sxs-lookup"><span data-stu-id="db5fa-p111">Light grays are most effective for gridlines. If you’re using Fabric, see the [Neutral Colors palette](https://developer.microsoft.com/fabric#/styles/colors).</span></span>

<span data-ttu-id="db5fa-145">На приведенном ниже рисунке показана визуализация данных с линиями сетки.</span><span class="sxs-lookup"><span data-stu-id="db5fa-145">The following image shows a data visualization with gridlines.</span></span>

![Изображение визуализации данных с линиями сетки](../images/data-visualization.png)

### <a name="legends"></a><span data-ttu-id="db5fa-147">Условные обозначения</span><span class="sxs-lookup"><span data-stu-id="db5fa-147">Legends</span></span>

<span data-ttu-id="db5fa-148">Условные обозначения необходимы для следующего:</span><span class="sxs-lookup"><span data-stu-id="db5fa-148">Add legends if necessary to:</span></span>

- <span data-ttu-id="db5fa-149">различения рядов данных;</span><span class="sxs-lookup"><span data-stu-id="db5fa-149">Distinguish between series</span></span>
- <span data-ttu-id="db5fa-150">представления изменений масштаба и значений.</span><span class="sxs-lookup"><span data-stu-id="db5fa-150">Present scale or value changes</span></span>

<span data-ttu-id="db5fa-p112">Убедитесь, что условные обозначения выделяют точки данных, а не отвлекают от них. Располагайте условные обозначения следующим образом:</span><span class="sxs-lookup"><span data-stu-id="db5fa-p112">Make sure that your legends enhance the data ink and do not compete with it. Place legends:</span></span>


- <span data-ttu-id="db5fa-153">С выравниванием по левому краю над областью представления данных по умолчанию, если все обозначения помещаются над диаграммой.</span><span class="sxs-lookup"><span data-stu-id="db5fa-153">Flush left above the plot area by default, if all legend items fit above the chart.</span></span>
- <span data-ttu-id="db5fa-154">Справа вверху в области представления данных, если все обозначения не помещаются над диаграммой. При необходимости можно разрешить прокрутку списка.</span><span class="sxs-lookup"><span data-stu-id="db5fa-154">On the upper right side of the plot area, if all legend items do not fit above the chart, and make it scrollable, if necessary.</span></span>

<span data-ttu-id="db5fa-p113">Для наглядности придайте маркерам условных обозначений форму, соответствующую типу диаграммы. Например, круглые маркеры подходят для точечных и пузырьковых диаграмм. Для графиков подходят маркеры в виде сегментов линий.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p113">To optimize for readability and accessibility, map legend markers to the relevant chart shape. For example, use circle legend markers for scatter plot and bubble chart legends. Use line segment legend markers for line charts.</span></span>

### <a name="data-labels-and-tooltips"></a><span data-ttu-id="db5fa-158">Подписи и подсказки данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-158">Data labels and tooltips</span></span>

<span data-ttu-id="db5fa-p114">Убедитесь, что в подписях и подсказках данных используются достаточно большие отступы и подходящие типы. Используйте алгоритмы, чтобы свести к минимуму наложения. Например, всплывающая подсказка может по умолчанию появляться справа от данных, если соответствующая точка не находится слишком близко к правому краю.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p114">Ensure that data labels and tooltips have adequate white space and type variation. Use algorithms to minimize occlusion and collision. For example, a tooltip might surface to the right of a data point by default, but surface to the left if right edges are detected.</span></span>

## <a name="design-principles"></a><span data-ttu-id="db5fa-162">Принципы оформления</span><span class="sxs-lookup"><span data-stu-id="db5fa-162">Design principles</span></span>

<span data-ttu-id="db5fa-163">Команда разработчиков Office составила приведенный ниже список принципов оформления, которым мы следуем при визуализации данных для набора продуктов Office.</span><span class="sxs-lookup"><span data-stu-id="db5fa-163">The Office Design team created the following set of design principles, which we use when designing new data visualizations for the Office product suite.</span></span>

### <a name="visual-design-principles"></a><span data-ttu-id="db5fa-164">Принципы визуального оформления</span><span class="sxs-lookup"><span data-stu-id="db5fa-164">Visual design principles</span></span>

- <span data-ttu-id="db5fa-p115">Визуализация должна точно и качественно передавать данные, чтобы их было легче понять. Выделяйте данные с помощью вспомогательных элементов только в той степени, которой требует контекст. Избегайте лишних украшений (теней, контуров и т. д.), ненужных элементов и искажения данных.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p115">Visualizations should honor and enhance the data, making it easy to understand. Highlight the data, adding supporting elements only as needed to provide context. Avoid unnecessary embellishments (drop shadows, outlines, etc), chart junk, or data distortion.</span></span>
- <span data-ttu-id="db5fa-p116">Визуализация должна вызывать интерес за счет наглядных зрительных образов. Используйте традиционные шаблоны взаимодействия, элементы управления и понятные реакции системы.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p116">Visualizations should encourage exploration by providing rich visual feedback. Use well-established interaction patterns, interface controls, and clear system feedback.</span></span>
- <span data-ttu-id="db5fa-p117">Применяйте проверенные временем принципы оформления. Следуйте традиционным принципам типографии и визуальной передачи, чтобы улучшить оформление, повысить удобочитаемость и точно передать смысл.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p117">Embody time-honored design principles. Use established typographic and visual communication design principles to enhance form, readability, and meaning.</span></span>

### <a name="interaction-design-principles"></a><span data-ttu-id="db5fa-172">Принципы взаимодействия</span><span class="sxs-lookup"><span data-stu-id="db5fa-172">Interaction design principles</span></span>

- <span data-ttu-id="db5fa-173">Диаграмма должна вызывать интерес.</span><span class="sxs-lookup"><span data-stu-id="db5fa-173">Design to allow for exploration.</span></span>
- <span data-ttu-id="db5fa-174">Обеспечьте непосредственное взаимодействие с объектами, позволяющее взглянуть на данные с новой стороны (например, сортировку путем перетаскивания).</span><span class="sxs-lookup"><span data-stu-id="db5fa-174">Allow for direct interactions with objects that reveal new insights (sorting via drag, for example).</span></span>
- <span data-ttu-id="db5fa-175">Используйте простые, непосредственные и привычные модели взаимодействия.</span><span class="sxs-lookup"><span data-stu-id="db5fa-175">Use simple, direct, familiar interaction models.</span></span>

<span data-ttu-id="db5fa-176">Дополнительные сведения о создании понятных интерактивных представлений данных см. в статье [Принципы и распространенные ошибки оформления интерфейса](https://uitraps.com/).</span><span class="sxs-lookup"><span data-stu-id="db5fa-176">For more information about how to design user-friendly interactive data visualizations, see [UI Tenets and Traps](https://uitraps.com/).</span></span>

### <a name="motion-design-principles"></a><span data-ttu-id="db5fa-177">Принципы динамического оформления</span><span class="sxs-lookup"><span data-stu-id="db5fa-177">Motion design principles</span></span>

<span data-ttu-id="db5fa-p118">Движение — результат воздействия. Визуальные элементы должны двигаться в одном направлении и с одинаковой скоростью. Это относится к следующему:</span><span class="sxs-lookup"><span data-stu-id="db5fa-p118">Motion follows stimulus. Visual elements should move in the same direction at the same rate. This applies to:</span></span>

- <span data-ttu-id="db5fa-181">созданию диаграмм;</span><span class="sxs-lookup"><span data-stu-id="db5fa-181">Chart creation</span></span>
- <span data-ttu-id="db5fa-182">изменению типа диаграммы;</span><span class="sxs-lookup"><span data-stu-id="db5fa-182">Transition from one chart type to another chart type</span></span>
- <span data-ttu-id="db5fa-183">фильтрам;</span><span class="sxs-lookup"><span data-stu-id="db5fa-183">Filtering</span></span>
- <span data-ttu-id="db5fa-184">сортировке;</span><span class="sxs-lookup"><span data-stu-id="db5fa-184">Sorting</span></span>
- <span data-ttu-id="db5fa-185">сложению и вычитанию данных;</span><span class="sxs-lookup"><span data-stu-id="db5fa-185">Adding or subtracting data</span></span>
- <span data-ttu-id="db5fa-186">объединению и сегментации данных;</span><span class="sxs-lookup"><span data-stu-id="db5fa-186">Brushing or slicing data</span></span>
- <span data-ttu-id="db5fa-187">изменению размера диаграммы;</span><span class="sxs-lookup"><span data-stu-id="db5fa-187">Resizing a chart</span></span>

<span data-ttu-id="db5fa-p119">созданию ощущения непринужденности. При создании анимации следуйте таким рекомендациям:</span><span class="sxs-lookup"><span data-stu-id="db5fa-p119">Create a perception of causality. When staging animations:</span></span>

- <span data-ttu-id="db5fa-190">Проектируйте элементы по одному.</span><span class="sxs-lookup"><span data-stu-id="db5fa-190">Stage one thing at a time.</span></span> 
- <span data-ttu-id="db5fa-191">Изменяйте оси, прежде чем менять точки данных.</span><span class="sxs-lookup"><span data-stu-id="db5fa-191">Stage changes to axes before changes to data ink.</span></span>
- <span data-ttu-id="db5fa-192">Если объекты двигаются в одном направлении и с одинаковой скоростью, обрабатывайте их как группу.</span><span class="sxs-lookup"><span data-stu-id="db5fa-192">Stage and animate objects as a group if they are moving at the same speed in the same direction.</span></span>
- <span data-ttu-id="db5fa-p120">Собирайте элементы в группы не более чем из 4–5 объектов. Пользователям сложно отслеживать более 4–5 независимых объектов.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p120">Stage data elements in groups of no more than 4-5 objects. Viewers have difficulty tracking more than 4-5 objects independently.</span></span>

<span data-ttu-id="db5fa-195">Движение добавляет осмысленность.</span><span class="sxs-lookup"><span data-stu-id="db5fa-195">Motion adds meaning.</span></span>

- <span data-ttu-id="db5fa-196">Анимация помогает пользователям ориентироваться в изменениях данных, создает контекст и заменяет комментарии.</span><span class="sxs-lookup"><span data-stu-id="db5fa-196">Animations increase user comprehension of changes to the data, provide context, and act as a non-verbal annotation layer.</span></span>
- <span data-ttu-id="db5fa-197">Движение должно происходить в понятном координатном пространстве визуализации.</span><span class="sxs-lookup"><span data-stu-id="db5fa-197">Motion should occur in a meaningful coordinate space of the visualization.</span></span>
- <span data-ttu-id="db5fa-198">Анимация должна соответствовать визуальному оформлению.</span><span class="sxs-lookup"><span data-stu-id="db5fa-198">Tailor the animation to the visual.</span></span> 
- <span data-ttu-id="db5fa-199">Не используйте анимацию без необходимости.</span><span class="sxs-lookup"><span data-stu-id="db5fa-199">Avoid gratuitous animations.</span></span>

<span data-ttu-id="db5fa-200">Движение следует за данными.</span><span class="sxs-lookup"><span data-stu-id="db5fa-200">Motion follows data.</span></span>

- <span data-ttu-id="db5fa-p121">Сохраняйте сопоставления данных. Если область привязана к показателю, сохраняйте ее при переходе.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p121">Preserve data mappings. If an area is tied to a measure, maintain that area in transition.</span></span>
- <span data-ttu-id="db5fa-p122">Все анимации должны быть выдержаны в одном стиле. По возможности согласуйте анимацию визуализации данных с оформлением Office. Используйте аналогичные анимации для похожих диаграмм.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p122">Maintain a consistent animation design language. Where possible, map data visualization animation to existing Office motion design language. Use similar animations for similar chart types.</span></span>

## <a name="accessibility-in-data-visualizations"></a><span data-ttu-id="db5fa-206">Специальные возможности для визуализации данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-206">Accessibility in data visualizations</span></span>

- <span data-ttu-id="db5fa-p123">Цвет не должен быть единственным способом передачи информации. В противном случае люди, страдающие дальтонизмом, не смогут толковать результаты. По мере возможности используйте для передачи информации не только цвет, но и форму, размер и текстуры.</span><span class="sxs-lookup"><span data-stu-id="db5fa-p123">Do not use color as the only way to communicate information. People who are color blind will not be able to interpret the results. Use shape, size and texture in addition to color when possible to communicate information.</span></span>
- <span data-ttu-id="db5fa-210">Обеспечьте возможность управлять с помощью клавиатуры всеми интерактивными элементами, такими как кнопки и списки.</span><span class="sxs-lookup"><span data-stu-id="db5fa-210">Make all interactive elements, such as push buttons or pick lists, accessible from a keyboard.</span></span>
- <span data-ttu-id="db5fa-211">Отправляйте события специальных возможностей средствам чтения с экрана для объявления об изменениях фокуса, всплывающих подсказках и т. д.</span><span class="sxs-lookup"><span data-stu-id="db5fa-211">Send accessibility events to screen readers to announce focus changes, tooltips, and so on.</span></span>

## <a name="see-also"></a><span data-ttu-id="db5fa-212">См. также</span><span class="sxs-lookup"><span data-stu-id="db5fa-212">See also</span></span> 

- [<span data-ttu-id="db5fa-213">Пять лучших библиотек для визуализации данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-213">The Five Best Libraries for Building Data Visualizations</span></span>](https://www.fastcompany.com/3029760/the-five-best-libraries-for-building-data-vizualizations)
- [<span data-ttu-id="db5fa-214">Визуальное представление количественных данных</span><span class="sxs-lookup"><span data-stu-id="db5fa-214">The Visual Display of Quantitative Information</span></span>](https://www.edwardtufte.com/tufte/books_vdqi)
