async function openSystemFileDialog({ title, purpose, mode = "file", startPath = "" }) {
  const response = await fetch("/api/system-file-dialog", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      title,
      purpose,
      mode,
      startPath,
    }),
  });
  const payload = await response.json();
  if (!response.ok) {
    throw new Error(payload.errors?.[0]?.message ?? `Could not open file picker (${response.status})`);
  }
  if (payload.canceled) {
    return "";
  }
  return String(payload.path ?? "");
}

export { openSystemFileDialog };
