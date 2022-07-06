using ADUserUpdate.API;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;


var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddScoped<IUserService, UserService>();
builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("AppSettings"));

var app = builder.Build();

app.MapGet("/", ([FromQuery(Name = "upn")] string upn, [FromServices] IUserService userService)
    => userService.GetUser(upn)
);

app.MapPost("/", ([FromServices] IUserService userService, [FromBody] User user)
    => userService.UpdateUser(user)
);

app.MapGet("/expiring", ([FromQuery(Name = "daysFromToday")] int daysFromToday, [FromServices] IUserService userService)
    => userService.GetExpiringUsers(daysFromToday)
);

app.MapGet("/expired", ([FromServices] IUserService userService)
    => userService.GetExpiredUsers()
);

app.MapPost("/disable", ([FromQuery(Name = "upn")] string upn, [FromServices] IUserService userService)
    => userService.DisableUser(upn)
);

app.Run();
