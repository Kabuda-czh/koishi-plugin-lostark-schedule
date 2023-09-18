export interface LostarkSchedule {
  id: number
  boss: string
  qq: string
  guildId: string
  name: string
  dps1: number
  dps2: number
  mercy: number
  lastDps1: number
  lastDps2: number
  lastMercy: number
  reason?: string
  days: { [key: string]: boolean }
  joinDate: string
  lastJoinDate: string
  uploadDate: string
}

export interface ScheduleUser {
  user: LostarkSchedule
  role: 'dps1' | 'dps2' | 'mercy'
}

export interface Round {
  users: ScheduleUser[]
}
