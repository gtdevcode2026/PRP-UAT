// Native OOXML column-chart generator. Pure string builders — no I/O, no zip.
// Turns a report's chart data (categories, series names/colors/values) into the
// xl/charts/chartN.xml + xl/drawings/drawingN.xml parts that Excel opens as a
// real, editable chart. Replaces the baked-PNG (renderStyledPng) embed.
// Dual-export: browser (window.NativeChart) + Node (module.exports) for tests.
(function (root) {
  'use strict';

  var C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  var A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  var R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function strRef(ref, cache) {
    var pts = cache.map(function (v, i) {
      return '<c:pt idx="' + i + '"><c:v>' + esc(v) + '</c:v></c:pt>';
    }).join('');
    return '<c:strRef><c:f>' + esc(ref) + '</c:f><c:strCache>' +
      '<c:ptCount val="' + cache.length + '"/>' + pts + '</c:strCache></c:strRef>';
  }

  function numRef(ref, cache) {
    var pts = cache.map(function (v, i) {
      var num = (v === null || v === undefined || v === '') ? 0 : v;
      return '<c:pt idx="' + i + '"><c:v>' + num + '</c:v></c:pt>';
    }).join('');
    return '<c:numRef><c:f>' + esc(ref) + '</c:f><c:numCache>' +
      '<c:formatCode>General</c:formatCode><c:ptCount val="' + cache.length + '"/>' +
      pts + '</c:numCache></c:numRef>';
  }

  function fill(color) {
    return '<c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>' +
      '<a:ln><a:noFill/></a:ln></c:spPr>';
  }

  function seriesXml(s, idx, cat, dLbls) {
    var tx = s.name.ref
      ? '<c:tx>' + strRef(s.name.ref, [s.name.lit]) + '</c:tx>'
      : '<c:tx><c:v>' + esc(s.name.lit) + '</c:v></c:tx>';
    var dPts = (s.points || []).map(function (p) {
      return '<c:dPt><c:idx val="' + p.idx + '"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>' +
        fill(p.color) + '</c:dPt>';
    }).join('');
    return '<c:ser><c:idx val="' + idx + '"/><c:order val="' + idx + '"/>' +
      tx + fill(s.color) + '<c:invertIfNegative val="0"/>' + dPts + dLbls +
      '<c:cat>' + strRef(cat.ref, cat.cache) + '</c:cat>' +
      '<c:val>' + numRef(s.values.ref, s.values.cache) + '</c:val></c:ser>';
  }

  function dataLabels(dl) {
    if (!dl) return '';
    var numFmt = dl.numFmt ? '<c:numFmt formatCode="' + dl.numFmt + '" sourceLinked="0"/>' : '';
    var txPr = dl.color
      ? '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:solidFill>' +
        '<a:srgbClr val="' + dl.color + '"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
      : '';
    return '<c:dLbls>' + numFmt + txPr + '<c:dLblPos val="' + dl.position + '"/>' +
      '<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/>' +
      '<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';
  }

  function axisTxt(color) {
    if (!color) return '';
    return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:solidFill>' +
      '<a:srgbClr val="' + color + '"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
  }

  function chartXml(def) {
    var CAT = 111111111, VAL = 222222222;
    var dl = dataLabels(def.dataLabels);
    var sers = def.series.map(function (s, i) { return seriesXml(s, i, def.categories, dl); }).join('');
    var overlap = def.grouping === 'stacked' ? '<c:overlap val="100"/>' : '<c:overlap val="-27"/>';
    var gap = '<c:gapWidth val="' + (def.grouping === 'stacked' ? 50 : 150) + '"/>';
    var titleRPr = def.axisColor
      ? '<a:rPr lang="en-US" b="1"><a:solidFill><a:srgbClr val="' + def.axisColor + '"/></a:solidFill></a:rPr>'
      : '<a:rPr lang="en-US" b="1"/>';
    var title = def.title
      ? '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>' + titleRPr + '<a:t>' + esc(def.title) +
        '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>'
      : '<c:autoTitleDeleted val="1"/>';
    var legend = def.legend
      ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/>' + axisTxt(def.axisColor) + '</c:legend>'
      : '';
    var chartSpaceFill = def.chartBg
      ? '<c:spPr><a:solidFill><a:srgbClr val="' + def.chartBg + '"/></a:solidFill></c:spPr>' : '';
    var plotFill = def.plotBg
      ? '<c:spPr><a:solidFill><a:srgbClr val="' + def.plotBg + '"/></a:solidFill></c:spPr>'
      : '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>';
    // barDir 'col' = vertical columns (default), 'bar' = horizontal bars.
    // For horizontal bars the category axis sits on the left, value axis below.
    var barDir = def.barDir || 'col';
    var catPos = barDir === 'bar' ? 'l' : 'b';
    var valPos = barDir === 'bar' ? 'b' : 'l';
    var catOrient = def.catReversed ? 'maxMin' : 'minMax';
    var valNumFmt = def.valNumFmt ? '<c:numFmt formatCode="' + def.valNumFmt + '" sourceLinked="0"/>' : '';
    // hideValAx mirrors xlsxwriter's set_y_axis({visible: False, major_gridlines:
    // {visible: False}}) — used where the bars carry their own data labels and a
    // second scale would only add noise. Default keeps the axis and gridlines.
    var valDelete = def.hideValAx ? '1' : '0';
    var valGridlines = def.hideValAx ? '' : '<c:majorGridlines/>';
    // catTickLblPos mirrors set_x_axis({label_position: ...}); 'low' pins the
    // category labels below the plot instead of floating them at the axis line.
    var catTickLblPos = def.catTickLblPos
      ? '<c:tickLblPos val="' + def.catTickLblPos + '"/>' : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<c:chartSpace xmlns:c="' + C + '" xmlns:a="' + A + '" xmlns:r="' + R + '">' +
      '<c:roundedCorners val="0"/><c:chart>' + title +
      '<c:plotArea><c:layout/>' +
      '<c:barChart><c:barDir val="' + barDir + '"/><c:grouping val="' + def.grouping + '"/>' +
      '<c:varyColors val="0"/>' + sers + gap + overlap +
      '<c:axId val="' + CAT + '"/><c:axId val="' + VAL + '"/></c:barChart>' +
      '<c:catAx><c:axId val="' + CAT + '"/><c:scaling><c:orientation val="' + catOrient + '"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="' + catPos + '"/>' + catTickLblPos + axisTxt(def.axisColor) +
      '<c:crossAx val="' + VAL + '"/></c:catAx>' +
      '<c:valAx><c:axId val="' + VAL + '"/><c:scaling><c:orientation val="minMax"/></c:scaling>' +
      '<c:delete val="' + valDelete + '"/><c:axPos val="' + valPos + '"/>' + valGridlines + valNumFmt + axisTxt(def.axisColor) +
      '<c:crossAx val="' + CAT + '"/></c:valAx>' +
      plotFill + '</c:plotArea>' + legend +
      '<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart>' +
      chartSpaceFill + '</c:chartSpace>';
  }

  function drawingXml(a, relId) {
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<xdr:wsDr xmlns:xdr="' + XDR + '" xmlns:a="' + A + '" xmlns:r="' + R + '" xmlns:c="' + C + '">' +
      '<xdr:twoCellAnchor>' +
      '<xdr:from><xdr:col>' + a.fromCol + '</xdr:col><xdr:colOff>0</xdr:colOff>' +
      '<xdr:row>' + a.fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + a.toCol + '</xdr:col><xdr:colOff>0</xdr:colOff>' +
      '<xdr:row>' + a.toRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>' +
      '<xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
      '<a:graphic><a:graphicData uri="' + C + '"><c:chart xmlns:c="' + C + '" r:id="' + relId + '"/>' +
      '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>' +
      '</xdr:twoCellAnchor></xdr:wsDr>';
  }

  var API = { chartXml: chartXml, drawingXml: drawingXml, _esc: esc };
  root.NativeChart = API;
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
