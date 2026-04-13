"use client";

import { motion } from "motion/react";
import { ReactNode } from "react";

interface Props {
  children: ReactNode;
  className?: string;
  delay?: number;
}

export default function Section({ children, className = "", delay = 0 }: Props) {
  return (
    <motion.section
      initial={{ opacity: 0, y: 24 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true, margin: "-80px" }}
      transition={{ duration: 0.6, delay, ease: "easeOut" }}
      className={className}
    >
      {children}
    </motion.section>
  );
}
