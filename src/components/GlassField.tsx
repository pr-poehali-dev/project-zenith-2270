import Icon from "@/components/ui/icon"

interface GlassFieldProps {
  label: string
  value: string
  onChange: (value: string) => void
  placeholder?: string
  icon?: string
  multiline?: boolean
  type?: string
}

export function GlassField({ label, value, onChange, placeholder, icon, multiline, type = "text" }: GlassFieldProps) {
  const sharedClass =
    "w-full rounded-2xl bg-white/60 px-4 py-3 text-[15px] text-gray-800 placeholder:text-gray-400 outline-none transition focus:bg-white/90"

  const sharedStyle = {
    backdropFilter: "blur(20px) saturate(180%)",
    WebkitBackdropFilter: "blur(20px) saturate(180%)",
    border: "1px solid rgba(255,255,255,0.6)",
    boxShadow: "inset 0 1px 1px rgba(255,255,255,0.9), 0 2px 8px rgba(0,0,0,0.04)",
  }

  return (
    <div className="space-y-1.5">
      <label className="flex items-center gap-1.5 text-[13px] font-medium text-gray-600 pl-1">
        {icon && <Icon name={icon} size={14} className="text-gray-500" />}
        {label}
      </label>
      {multiline ? (
        <textarea
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          rows={4}
          className={`${sharedClass} resize-none`}
          style={sharedStyle}
        />
      ) : (
        <input
          type={type}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={placeholder}
          className={sharedClass}
          style={sharedStyle}
        />
      )}
    </div>
  )
}
