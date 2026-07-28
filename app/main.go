package main

import (
	"embed"
	"log"
	"path/filepath"

	"github.com/wailsapp/wails/v2"
	"github.com/wailsapp/wails/v2/pkg/options"
	"github.com/wailsapp/wails/v2/pkg/options/assetserver"

	"swissknife-app/internal/auditlog"
	"swissknife-app/internal/journal"
	"swissknife-app/internal/secrets"
	"swissknife-app/internal/services"
	"swissknife-app/internal/session"
)

//go:embed all:frontend/dist
var assets embed.FS

func main() {
	store, err := secrets.NewStore()
	if err != nil {
		log.Fatal("init config dir: ", err)
	}
	audit := auditlog.New(store.Dir())
	sess := session.New(audit)
	runs := journal.New(filepath.Join(store.Dir(), "runs"))
	runs.Prune(200) // keep the newest 200 runs
	sess.SetJournal(runs)
	sess.SetConfigDir(store.Dir())

	updateSvc := services.NewUpdateService(resolveVersion())
	app := NewApp(sess, updateSvc)

	err = wails.Run(&options.App{
		Title:  "SwissKnife for MS Graph",
		Width:  1280,
		Height: 800,
		AssetServer: &assetserver.Options{
			Assets: assets,
		},
		BackgroundColour: &options.RGBA{R: 15, G: 17, B: 23, A: 1},
		OnStartup:        app.startup,
		Bind: []interface{}{
			app,
			services.NewConnectService(sess, store),
			services.NewDashboardService(sess),
			services.NewAccessService(sess),
			services.NewPlaybookService(sess),
			services.NewMirrorService(sess),
			services.NewUsersService(sess),
			services.NewAuthMethodsService(sess),
			services.NewRolesService(sess),
			services.NewLicensingService(sess),
			services.NewGroupsService(sess),
			services.NewTeamsService(sess),
			services.NewChatsService(sess),
			services.NewMailService(sess),
			services.NewCalendarService(sess),
			services.NewDriveService(sess),
			services.NewIntuneService(sess),
			services.NewDevicesService(sess),
			services.NewAppsService(sess),
			services.NewSecurityService(sess),
			services.NewReportsService(sess),
			services.NewCleanupService(sess),
			services.NewServiceHealthService(sess),
			services.NewAuditService(sess),
			services.NewRawService(sess),
			services.NewJournalService(sess),
			services.NewNotifyService(sess),
			updateSvc,
		},
	})

	if err != nil {
		log.Fatal(err)
	}
}
