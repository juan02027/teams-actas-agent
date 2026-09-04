import { useAuth } from "@/_core/hooks/useAuth";
import { Avatar, AvatarFallback } from "@/components/ui/avatar";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarHeader,
  SidebarInset,
  SidebarMenu,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarProvider,
  SidebarTrigger,
  useSidebar,
} from "@/components/ui/sidebar";
import { startLogin } from "@/const";
import { useIsMobile } from "@/hooks/useMobile";
import { CalendarDays, ClipboardList, FileText, LayoutDashboard, ListChecks, LogOut, PanelLeft, Settings2, Video } from "lucide-react";
import { CSSProperties, useEffect, useRef, useState } from "react";
import { useLocation } from "wouter";
import { DashboardLayoutSkeleton } from "./DashboardLayoutSkeleton";
import { Button } from "./ui/button";

const menuItems = [
  { icon: LayoutDashboard, label: "Centro de control", path: "/" },
  { icon: CalendarDays, label: "Reuniones", path: "/reuniones" },
  { icon: FileText, label: "Actas", path: "/actas" },
  { icon: Video, label: "Grabaciones", path: "/grabaciones" },
  { icon: ListChecks, label: "Compromisos", path: "/compromisos" },
  { icon: Settings2, label: "Configuración", path: "/configuracion" },
];

const SIDEBAR_WIDTH_KEY = "sidebar-width";
const DEFAULT_WIDTH = 280;
const MIN_WIDTH = 200;
const MAX_WIDTH = 480;

export default function DashboardLayout({ children }: { children: React.ReactNode }) {
  const [sidebarWidth, setSidebarWidth] = useState(() => {
    const saved = localStorage.getItem(SIDEBAR_WIDTH_KEY);
    return saved ? parseInt(saved, 10) : DEFAULT_WIDTH;
  });
  const { loading, user } = useAuth();
  const localDevelopmentMode = import.meta.env.DEV || import.meta.env.VITE_LOCAL_DEMO_MODE === "true";
  useEffect(() => {
    localStorage.setItem(SIDEBAR_WIDTH_KEY, sidebarWidth.toString());
  }, [sidebarWidth]);

  if (loading) return <DashboardLayoutSkeleton />;

  if (!user && !localDevelopmentMode) {
    return (
      <div className="min-h-screen bg-[#f4f7f7] flex items-center justify-center px-6">
        <div className="w-full max-w-md rounded-3xl bg-white p-10 shadow-[0_24px_70px_rgba(15,37,48,0.12)] text-center">
          <div className="mx-auto mb-6 flex h-14 w-14 items-center justify-center rounded-2xl bg-[#dff3ee] text-[#0d776c]">
            <ClipboardList className="h-7 w-7" />
          </div>
          <p className="mb-3 text-xs font-semibold uppercase tracking-[0.2em] text-[#0d776c]">Teams Actas Agent</p>
          <h1 className="text-3xl font-semibold tracking-tight text-[#102c36]">Panel privado del operador</h1>
          <p className="mt-4 text-sm leading-6 text-[#67808a]">Inicia sesión para controlar qué reuniones se graban, cuáles se procesan y cuándo se distribuyen sus documentos.</p>
          <Button onClick={() => startLogin()} size="lg" className="mt-8 w-full bg-[#0d776c] text-white shadow-lg shadow-[#0d776c]/20 hover:bg-[#095f57]">Entrar al panel</Button>
        </div>
      </div>
    );
  }

  return (
    <SidebarProvider style={{ "--sidebar-width": `${sidebarWidth}px` } as CSSProperties}>
      <DashboardLayoutContent setSidebarWidth={setSidebarWidth}>{children}</DashboardLayoutContent>
    </SidebarProvider>
  );
}

type DashboardLayoutContentProps = { children: React.ReactNode; setSidebarWidth: (width: number) => void };

function DashboardLayoutContent({ children, setSidebarWidth }: DashboardLayoutContentProps) {
  const { user, logout } = useAuth();
  const [location, setLocation] = useLocation();
  const { state, toggleSidebar } = useSidebar();
  const isCollapsed = state === "collapsed";
  const [isResizing, setIsResizing] = useState(false);
  const sidebarRef = useRef<HTMLDivElement>(null);
  const activeMenuItem = menuItems.find((item) => item.path === location);
  const isMobile = useIsMobile();
  const [m365Profile, setM365Profile] = useState<{ name: string; email: string } | null>(() => {
    try { return JSON.parse(localStorage.getItem("m365-profile") || "null") as { name: string; email: string } | null; } catch { return null; }
  });
  const displayName = m365Profile?.name || user?.name || "Operador local";
  const displayEmail = m365Profile?.email || "Acceso administrador";

  useEffect(() => {
    const updateProfile = () => {
      try { setM365Profile(JSON.parse(localStorage.getItem("m365-profile") || "null") as { name: string; email: string } | null); } catch { setM365Profile(null); }
    };
    window.addEventListener("m365-profile-updated", updateProfile);
    return () => window.removeEventListener("m365-profile-updated", updateProfile);
  }, []);

  useEffect(() => {
    if (isCollapsed) setIsResizing(false);
  }, [isCollapsed]);

  useEffect(() => {
    const handleMouseMove = (event: MouseEvent) => {
      if (!isResizing) return;
      const sidebarLeft = sidebarRef.current?.getBoundingClientRect().left ?? 0;
      const newWidth = event.clientX - sidebarLeft;
      if (newWidth >= MIN_WIDTH && newWidth <= MAX_WIDTH) setSidebarWidth(newWidth);
    };
    const handleMouseUp = () => setIsResizing(false);
    if (isResizing) {
      document.addEventListener("mousemove", handleMouseMove);
      document.addEventListener("mouseup", handleMouseUp);
      document.body.style.cursor = "col-resize";
      document.body.style.userSelect = "none";
    }
    return () => {
      document.removeEventListener("mousemove", handleMouseMove);
      document.removeEventListener("mouseup", handleMouseUp);
      document.body.style.cursor = "";
      document.body.style.userSelect = "";
    };
  }, [isResizing, setSidebarWidth]);

  return (
    <>
      <div className="relative" ref={sidebarRef}>
        <Sidebar collapsible="icon" className="border-r border-[#dbe8e5] bg-[#f8fbfa] text-[#102c36] shadow-[10px_0_35px_rgba(16,44,54,0.06)]" disableTransition={isResizing}>
          <SidebarHeader className="h-[86px] justify-center border-b border-[#dbe8e5]">
            <div className="flex w-full items-center gap-3 px-2">
              <button onClick={toggleSidebar} className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl bg-[#e2f1ed] text-[#102c36] transition hover:bg-[#cfe7e0] focus:outline-none focus-visible:ring-2 focus-visible:ring-[#5d887f]" aria-label="Contraer navegación">
                <PanelLeft className="h-4 w-4" />
              </button>
              {!isCollapsed && <div className="min-w-0"><p className="truncate text-sm font-bold tracking-tight text-[#102c36]">Teams Actas Agent</p><p className="mt-1 truncate text-[10px] font-semibold uppercase tracking-[0.18em] text-[#304b50]">Operaciones M365</p></div>}
            </div>
          </SidebarHeader>
          <SidebarContent className="gap-0 px-2 py-4">
            <p className="mb-3 px-3 text-[10px] font-bold uppercase tracking-[0.2em] text-[#102c36] group-data-[collapsible=icon]:hidden">Navegación</p>
            <SidebarMenu>
              {menuItems.map((item) => {
                const isActive = location === item.path;
                return <SidebarMenuItem key={item.path}><SidebarMenuButton isActive={isActive} onClick={() => setLocation(item.path)} tooltip={item.label} className={`h-11 rounded-xl font-semibold transition ${isActive ? "bg-[#d9ebe5] text-[#102c36] shadow-sm hover:bg-[#cce3dc]" : "text-[#102c36] hover:bg-[#e7f2ef]"}`}><item.icon className="h-4 w-4 text-[#102c36]" /><span>{item.label}</span></SidebarMenuButton></SidebarMenuItem>;
              })}
            </SidebarMenu>
          </SidebarContent>
          <SidebarFooter className="border-t border-[#dbe8e5] p-3">
            <DropdownMenu>
              <DropdownMenuTrigger asChild><button className="flex w-full items-center gap-3 rounded-xl px-1 py-2 text-left transition hover:bg-[#e7f2ef] focus:outline-none focus-visible:ring-2 focus-visible:ring-[#5d887f] group-data-[collapsible=icon]:justify-center"><Avatar className="h-9 w-9 shrink-0 border border-[#b9d7d0] bg-[#d9ebe5]"><AvatarFallback className="bg-[#d9ebe5] text-xs font-bold text-[#102c36]">{displayName.charAt(0).toUpperCase()}</AvatarFallback></Avatar><div className="min-w-0 flex-1 group-data-[collapsible=icon]:hidden"><p className="truncate text-sm font-bold text-[#102c36]">{displayName}</p><p className="mt-1 truncate text-xs font-semibold text-[#304b50]">{displayEmail}</p></div></button></DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="w-48"><DropdownMenuItem onClick={logout} className="cursor-pointer text-destructive focus:text-destructive"><LogOut className="mr-2 h-4 w-4" />Cerrar sesión</DropdownMenuItem></DropdownMenuContent>
            </DropdownMenu>
          </SidebarFooter>
        </Sidebar>
        <div className={`absolute right-0 top-0 z-50 h-full w-1 cursor-col-resize transition-colors hover:bg-[#0d776c]/30 ${isCollapsed ? "hidden" : ""}`} onMouseDown={() => !isCollapsed && setIsResizing(true)} />
      </div>
      <SidebarInset className="bg-[#f4f7f7]">
        {isMobile && <div className="sticky top-0 z-40 flex h-14 items-center gap-2 border-b bg-white/95 px-2 backdrop-blur"><SidebarTrigger className="h-9 w-9 rounded-lg" /><span className="text-sm font-medium text-[#102c36]">{activeMenuItem?.label ?? "Panel"}</span></div>}
        <main className="min-h-screen flex-1">{children}</main>
      </SidebarInset>
    </>
  );
}
