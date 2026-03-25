import { useEffect, useRef } from "react";
import { renderPptx } from "pptx-react-renderer";

export function ReactBasicExample({ file }: { file: ArrayBuffer | null }) {
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!file || !containerRef.current) return;

    void renderPptx(file, {
      container: containerRef.current,
      scale: 0.65,
      showSlideNumbers: true,
      theme: "light",
    });
  }, [file]);

  return <div ref={containerRef} />;
}
