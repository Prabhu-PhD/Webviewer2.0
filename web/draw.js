export function createDrawController(ui) {
  let ctx, isDrawing = false;
  let lastX = 0, lastY = 0;
  let currentColor = "#e05252";
  let strokes = [];        // [ { color, points: [{x,y}] }, ... ]
  let currentStroke = null;

  function setupDrawing() {
    ctx = ui.drawCanvas.getContext("2d");

    ui.drawCanvas.addEventListener("mousedown", startDrawing);
    ui.drawCanvas.addEventListener("mousemove", draw);
    ui.drawCanvas.addEventListener("mouseup", stopDrawing);
    ui.drawCanvas.addEventListener("mouseout", stopDrawing);

    ui.drawCanvas.addEventListener("touchstart", handleTouchStart, { passive: false });
    ui.drawCanvas.addEventListener("touchmove", handleTouchMove, { passive: false });
    ui.drawCanvas.addEventListener("touchend", stopDrawing);
  }

  function resizeCanvas() {
    if (!ui.drawCanvas) return;
    const rect = ui.shell.getBoundingClientRect();
    ui.drawCanvas.width = rect.width;
    ui.drawCanvas.height = rect.height;
    _applyCtxDefaults();
    _redraw();
  }

  function _applyCtxDefaults() {
    ctx.lineWidth = 4;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
  }

  function _redraw() {
    ctx.clearRect(0, 0, ui.drawCanvas.width, ui.drawCanvas.height);
    for (const stroke of strokes) {
      if (stroke.points.length < 2) continue;
      ctx.beginPath();
      ctx.strokeStyle = stroke.color;
      ctx.moveTo(stroke.points[0].x, stroke.points[0].y);
      for (let i = 1; i < stroke.points.length; i++) {
        ctx.lineTo(stroke.points[i].x, stroke.points[i].y);
      }
      ctx.stroke();
    }
  }

  function getPos(e) {
    const rect = ui.drawCanvas.getBoundingClientRect();
    const touch = e.touches?.[0];
    const clientX = touch ? touch.clientX : e.clientX;
    const clientY = touch ? touch.clientY : e.clientY;
    return { x: clientX - rect.left, y: clientY - rect.top };
  }

  function startDrawing(e) {
    isDrawing = true;
    const pos = getPos(e);
    lastX = pos.x;
    lastY = pos.y;
    currentStroke = { color: currentColor, points: [{ x: lastX, y: lastY }] };
  }

  function draw(e) {
    if (!isDrawing || !currentStroke) return;
    e.preventDefault();
    const pos = getPos(e);
    ctx.beginPath();
    ctx.strokeStyle = currentColor;
    ctx.moveTo(lastX, lastY);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
    currentStroke.points.push({ x: pos.x, y: pos.y });
    lastX = pos.x;
    lastY = pos.y;
  }

  function handleTouchStart(e) {
    e.preventDefault();
    startDrawing(e);
  }

  function handleTouchMove(e) {
    e.preventDefault();
    draw(e);
  }

  function stopDrawing() {
    if (!isDrawing) return;
    isDrawing = false;
    if (currentStroke && currentStroke.points.length > 1) {
      strokes.push(currentStroke);
    }
    currentStroke = null;
  }

  function clearCanvas() {
    if (ctx && ui.drawCanvas) {
      strokes = [];
      currentStroke = null;
      ctx.clearRect(0, 0, ui.drawCanvas.width, ui.drawCanvas.height);
    }
  }

  function undoStroke() {
    if (strokes.length === 0) return;
    strokes.pop();
    _applyCtxDefaults();
    _redraw();
  }

  function setColor(color) {
    currentColor = color;
    _applyCtxDefaults();
  }

  return { setupDrawing, resizeCanvas, clearCanvas, undoStroke, setColor };
}
