using MyMicroservice;
var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();

    builder.Services.AddSingleton<IMyService, MyService>();
    builder.Services.AddSingleton<IValidationService, ValidationService>();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
    
        app.UseDeveloperExceptionPage();
    
}

app.UseRouting();

    app.UseEndpoints(endpoints =>
    {
        endpoints.MapControllers();
    });

app.UseHttpsRedirection();

// app.UseMiddleware<CustomAuthorizationMiddleware>();

app.UseAuthorization();

app.MapControllers();


app.Run();
