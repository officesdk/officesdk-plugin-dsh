// The working selection is painted inside the editor iframe, the review card in the host.
export const WORK_BORDER_STOPS = [['0%', '#70adff'], ['50%', '#cb8ced'], ['100%', '#60cbd2']];
export const WORK_BORDER_CSS = 'linear-gradient(to right, ' +
  WORK_BORDER_STOPS.map(([offset, color]) => color + ' ' + offset).join(', ') + ')';
