/** How text is measured by the engine's estimates (the real renderer is a little wider than this). */

export const CM_TO_PT = 28.3465;
export const INSET_CM = 0.25;

/** How wide a text is, in em: Chinese characters are one em wide, Latin ones about half. */
export const widthEm = (text: string) => [...text].reduce((sum, char) => sum + (/[⺀-鿿＀-￯]/.test(char) ? 1 : 0.55), 0);
