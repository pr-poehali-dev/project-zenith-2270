import { icons, LucideProps } from "lucide-react"

interface IconProps extends LucideProps {
  name: string
  fallback?: string
}

const Icon = ({ name, fallback = "CircleAlert", ...props }: IconProps) => {
  const LucideIcon = (icons as Record<string, React.ComponentType<LucideProps>>)[name] ||
    (icons as Record<string, React.ComponentType<LucideProps>>)[fallback]

  return <LucideIcon {...props} />
}

export default Icon
