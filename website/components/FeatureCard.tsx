"use client";

import { motion } from "motion/react";
import { LucideIcon } from "lucide-react";

interface Props {
  icon: LucideIcon;
  title: string;
  children: React.ReactNode;
  delay?: number;
}

export default function FeatureCard({ icon: Icon, title, children, delay = 0 }: Props) {
  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      transition={{ duration: 0.5, delay }}
      className="relative rounded-xl border border-line bg-panel/60 p-6 hover:border-brand-500/40 transition-colors"
    >
      <div className="w-10 h-10 rounded-lg bg-gradient-to-br from-brand-500/20 to-accent-500/20 border border-brand-500/30 flex items-center justify-center mb-4">
        <Icon size={20} className="text-brand-600 dark:text-brand-300" />
      </div>
      <h3 className="font-semibold text-fg mb-2">{title}</h3>
      <p className="text-sm text-fg-subtle leading-relaxed">{children}</p>
    </motion.div>
  );
}
