/**
 * React binding without a React dependency: pass your React in, get a component back.
 *
 *   import * as React from 'react';
 *   import { createDocxViewerComponent } from 'docxkeun/react';
 *   const DocxView = createDocxViewerComponent(React);
 *   <DocxView file={bytes} fonts={registry} onReady={(viewer) => ...} />
 */
import type { OpenOptions } from '../model/document.js';
import { DocxViewer, type ViewerOptions } from '../render/viewer.js';

/** The slice of React this binding uses. */
export interface ReactLike {
  createElement(type: string, props: Record<string, unknown> | null): unknown;
  useRef<T>(initial: T | null): { current: T | null };
  useEffect(effect: () => void | (() => void), deps?: unknown[]): void;
}

export interface DocxViewProps extends ViewerOptions {
  /** Document bytes (or a Blob / File). */
  file: Uint8Array | ArrayBuffer | Blob | null | undefined;
  openOptions?: OpenOptions;
  className?: string;
  style?: Record<string, string | number>;
  onReady?(viewer: DocxViewer): void;
  onError?(error: Error): void;
}

export function createDocxViewerComponent(React: ReactLike) {
  return function DocxView(props: DocxViewProps) {
    const host = React.useRef<HTMLDivElement>(null);
    const viewerRef = React.useRef<DocxViewer>(null);
    const { file, openOptions, className, style, onReady, onError, ...options } = props;

    React.useEffect(() => {
      if (!host.current) return;
      const viewer = new DocxViewer(host.current, options);
      viewerRef.current = viewer;
      return () => {
        viewer.destroy();
        viewerRef.current = null;
      };
      // Options are read once; remount (change `key`) to apply new ones.
    }, []);

    React.useEffect(() => {
      const viewer = viewerRef.current;
      if (!viewer || !file) return;
      let cancelled = false;
      viewer
        .load(file, openOptions)
        .then(() => {
          if (!cancelled) onReady?.(viewer);
        })
        .catch((e: unknown) => {
          if (!cancelled) onError?.(e instanceof Error ? e : new Error(String(e)));
        });
      return () => {
        cancelled = true;
      };
    }, [file, openOptions?.password]);

    return React.createElement('div', { ref: host, className, style: { height: '100%', ...style } });
  };
}
