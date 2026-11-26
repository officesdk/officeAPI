// 视频帧：每个场景的每一步一页，最后一页是 README 里的视频封面
await PowerPoint.run(async (context) => {
  const plan = await renderScenes(context, VIDEO, true);
  const [poster] = await blankSlides(context, 1);
  SCENE_POSTER.draw(poster);
  await context.sync();
  return plan;
});
