import DocsSidebar from "@/components/DocsSidebar";

export default function DocsLayout({ children }: { children: React.ReactNode }) {
  return (
    <div className="max-w-6xl mx-auto px-6 py-12">
      <div className="flex gap-12">
        <DocsSidebar />
        <article className="flex-1 min-w-0 prose-invert-custom max-w-3xl">{children}</article>
      </div>
    </div>
  );
}
