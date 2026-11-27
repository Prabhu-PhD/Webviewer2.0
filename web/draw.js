export function createDrawController(ui) {
  let ctx, isDrawing = false, lastX = 0, lastY = 0;

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
    ctx.strokeStyle = "#e05252";
    ctx.lineWidth = 4;
    ctx.lineCap = "round";
    ctx.lineJoin = "round";
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
  }

  function draw(e) {
    if (!isDrawing) return;
    e.preventDefault();
    const pos = getPos(e);
    ctx.beginPath();
    ctx.moveTo(lastX, lastY);
    ctx.lineTo(pos.x, pos.y);
    ctx.stroke();
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
    isDrawing = false;
  }

  function clearCanvas() {
    if (ctx && ui.drawCanvas) {
      ctx.clearRect(0, 0, ui.drawCanvas.width, ui.drawCanvas.height);
    }
  }

  return { setupDrawing, resizeCanvas, clearCanvas };
}
