/**
 * Theme styles for PPTX renderer
 */

export interface Theme {
  background: string;
  slideBackground: string;
  text: string;
  border: string;
  accent: string;
  shadow: string;
}

export const lightTheme: Theme = {
  background: "#f5f5f5",
  slideBackground: "#ffffff",
  text: "#333333",
  border: "#e0e0e0",
  accent: "#2563eb",
  shadow: "0 4px 20px rgba(0,0,0,0.1)",
};

export const darkTheme: Theme = {
  background: "#1a1a1a",
  slideBackground: "#2a2a2a",
  text: "#ffffff",
  border: "#404040",
  accent: "#3b82f6",
  shadow: "0 4px 20px rgba(0,0,0,0.5)",
};

export function getThemeStyles(_theme: "light" | "dark"): string {
  return "";
}
